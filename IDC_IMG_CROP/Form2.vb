Imports System.Drawing
Imports Cyotek.Windows.Forms
'Cyotek.Windows.Forms.ImageBox

Imports System.Drawing.Drawing2D
Imports System.IO
Imports AForge.Video
Imports AForge.Video.DirectShow
Imports System.ComponentModel
Imports System.Drawing.Imaging

Public Class Form2
    Private _originalImage As Image ' Hình chưa xoay
    Private _currentAngle As Single = 0 ' Góc xoay hiện tại
    Dim cameras As FilterInfoCollection ' Danh sách camera
    Dim cam As VideoCaptureDevice ' Camera đang chọn
    Dim _videoDevices As FilterInfoCollection ' Danh sách thiết bị
    Dim _videoSource As VideoCaptureDevice ' Thiết bị đang chạy

    Dim SelectionRegion As New RectangleF(10, 10, 100, 100) ' Tọa độ mặc định

    ' 1. LOAD HÌNH
    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        ImageBox1.SelectionRegion = Rectangle.Empty
        ImageBox1.Invalidate()
        ImageBox2.Image = Nothing
        ImageBox3.Image = Nothing
        Try
            If _videoSource IsNot Nothing AndAlso _videoSource.IsRunning Then
                ' Ngừng nhận khung hình mới ngay lập tức
                RemoveHandler _videoSource.NewFrame, AddressOf VideoSource_NewFrame

                ' Ra lệnh dừng camera
                _videoSource.SignalToStop()

                ' Chờ một chút để camera kịp đóng hẳn (tránh lỗi giải phóng bộ nhớ đột ngột)
                _videoSource.WaitForStop()
                _videoSource = Nothing
                ImageBox1.Image = Nothing
                btnStart.Text = "Lấy từ Camera | OFF"
            End If
        Catch ex As Exception
            ' Bỏ qua lỗi nếu có khi đóng
        End Try

        Try
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Image Files|*.jpg;*.png;*.bmp;*.gif"

                If ofd.ShowDialog() = DialogResult.OK Then
                    ' Giải phóng ảnh gốc và ảnh hiển thị cũ nếu có
                    SafeDispose(_originalImage)
                    SafeDispose(ImageBox1.Image)

                    ' Load ảnh vào _originalImage từ file stream (để không khóa file)
                    Using fs As New FileStream(ofd.FileName, FileMode.Open, FileAccess.Read)
                        _originalImage = Image.FromStream(fs)
                    End Using

                    _currentAngle = 0
                    ImageBox1.Image = CType(_originalImage.Clone(), Image)
                    ImageBox1.SizeMode = ImageBoxSizeMode.Normal ' Hoặc Fit tùy thư viện bạn dùng

                    ' Tính Zoom vừa khung hình
                    Dim zoomW As Double = ImageBox1.ClientSize.Width / _originalImage.Width
                    Dim zoomH As Double = ImageBox1.ClientSize.Height / _originalImage.Height
                    Dim fitZoom As Integer = CInt(Math.Floor(Math.Min(zoomW, zoomH) * 100))

                    ImageBox1.Zoom = fitZoom
                    ImageBox1.AutoCenter = True

                End If
            End Using
        Catch ex As Exception

        End Try
    End Sub
    ' Hàm an toàn giải phóng ảnh
    Private Sub SafeDispose(ByRef img As Image)
        If img IsNot Nothing Then
            Try
                img.Dispose()
            Catch
            End Try
            img = Nothing
        End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 1. Tìm tất cả Camera hiện có trên máy
        _videoDevices = New FilterInfoCollection(FilterCategory.VideoInputDevice)

        If _videoDevices.Count > 0 Then
            ' Đưa tên các Camera vào ComboBox
            For Each device As FilterInfo In _videoDevices
                cbCameras.Items.Add(device.Name)
            Next
            cbCameras.SelectedIndex = 0 ' Chọn cái đầu tiên mặc định
        Else
            MessageBox.Show("Không tìm thấy Camera nào!")
        End If

        ' Cấu hình ImageBox1 cho việc chọn vùng (Crop)
        ImageBox1.SelectionMode = Cyotek.Windows.Forms.ImageBoxSelectionMode.Rectangle
        '
        Me.ImageBox3.BringToFront()
        Me.ImageBox2.BringToFront()

        Set_Text_Name()
    End Sub

#Region "Ham xu ly anh"

    ' SỰ KIỆN CẬP NHẬT HÌNH ẢNH TRỰC TIẾP
    Private Sub VideoSource_NewFrame(sender As Object, eventArgs As NewFrameEventArgs)
        Try
            ' 1. Kiểm tra nếu Form hoặc ImageBox đã bị hủy thì thoát ngay
            If Me.IsDisposed OrElse ImageBox1.IsDisposed Then Return

            ' Clone hình
            Dim bmp As Bitmap = DirectCast(eventArgs.Frame.Clone(), Bitmap)

            ' 2. Sử dụng BeginInvoke để an toàn hơn khi đóng App
            If ImageBox1.InvokeRequired Then
                ImageBox1.BeginInvoke(Sub()
                                          Try
                                              ' Kiểm tra lại lần nữa bên trong luồng UI
                                              If Not ImageBox1.IsDisposed Then
                                                  Dim oldImg = ImageBox1.Image
                                                  ImageBox1.Image = bmp
                                                  If oldImg IsNot Nothing Then oldImg.Dispose()
                                              Else
                                                  bmp.Dispose() ' Hủy bitmap nếu không gán được
                                              End If
                                          Catch
                                              bmp.Dispose()
                                          End Try
                                      End Sub)
            Else
                Dim oldImg = ImageBox1.Image
                ImageBox1.Image = bmp
                If oldImg IsNot Nothing Then oldImg.Dispose()
            End If
        Catch ex As Exception
            ' Tránh crash nếu có lỗi bất ngờ trong luồng camera
        End Try
    End Sub

#Region "Quay hinh"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        _currentAngle -= 0.5
        If _currentAngle < -45 Then _currentAngle = -45
        UpdateImageBox()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        _currentAngle += 0.5
        If _currentAngle > 45 Then _currentAngle = 45
        UpdateImageBox() ' Gọi Invalidate để ImageBox vẽ lại theo góc mới
    End Sub
#End Region
    ' Hàm cập nhật ảnh xoay
    Private Sub UpdateImageBox()
        If _originalImage IsNot Nothing Then
            ' Giải phóng ảnh hiện tại trong ImageBox nếu khác ảnh gốc
            If ImageBox1.Image IsNot Nothing AndAlso ImageBox1.Image IsNot _originalImage Then
                SafeDispose(ImageBox1.Image)
            End If

            ' Lấy ảnh đã xoay
            Dim rotated = GetRotatedImage(_originalImage, _currentAngle)
            ImageBox1.Image = rotated
        End If
    End Sub


    ' Bạn cần thêm sự kiện Paint cho ImageBox1 (hoặc PictureBox1)


    Public Function GetRotatedImage(img As Image, angle As Single) As Bitmap
        If img Is Nothing Then Return Nothing

        ' Tính toán góc và kích thước
        Dim angleRad As Double = (angle Mod 360) * Math.PI / 180
        Dim cosA As Double = Math.Abs(Math.Cos(angleRad))
        Dim sinA As Double = Math.Abs(Math.Sin(angleRad))

        Dim newWidth As Integer = CInt(Math.Max(1, Math.Ceiling(img.Width * cosA + img.Height * sinA)))
        Dim newHeight As Integer = CInt(Math.Max(1, Math.Ceiling(img.Width * sinA + img.Height * cosA)))

        ' Tạo Bitmap mới
        Dim bmp As New Bitmap(newWidth, newHeight)
        bmp.SetResolution(img.HorizontalResolution, img.VerticalResolution)

        Using g As Graphics = Graphics.FromImage(bmp)
            ' Chế độ vẽ ổn định nhất cho GDI+
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            g.Clear(Color.Transparent)

            ' Ma trận dịch chuyển và xoay
            g.TranslateTransform(newWidth / 2.0F, newHeight / 2.0F)
            g.RotateTransform(angle)

            ' Dùng DrawImage bản đơn giản nhất để tránh lỗi ExternalException
            g.DrawImage(img, New Rectangle(-img.Width / 2, -img.Height / 2, img.Width, img.Height))
        End Using

        Return bmp
    End Function





    Private Function ResizeImage(ByVal mg As Image, ByVal newWidth As Integer) As Image
        ' Tính toán chiều cao theo tỷ lệ của hình gốc
        Dim ratio As Double = newWidth / mg.Width
        Dim newHeight As Integer = CInt(mg.Height * ratio)

        ' Tạo một Bitmap mới với kích thước đã tính
        Dim bp As New Bitmap(newWidth, newHeight)
        Using g As Graphics = Graphics.FromImage(bp)
            ' Thiết lập chất lượng cao cho hình ảnh
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.DrawImage(mg, 0, 0, newWidth, newHeight)
        End Using
        Return bp
    End Function
    Private Sub UpdateSelection()
        ' Gán vùng chọn mới cho ImageBox1
        ImageBox1.SelectionRegion = SelectionRegion
        ' Ép ImageBox1 vẽ lại để thấy thay đổi ngay lập tức
        ImageBox1.Invalidate()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        SelectionRegion = ImageBox1.SelectionRegion
        SelectionRegion.Y -= 5
        UpdateSelection()
    End Sub

#End Region
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        ImageBox1.SelectionRegion = Rectangle.Empty
        ' Cập nhật lại hiển thị để mất khung màu đỏ ngay lập tức
        ImageBox1.Invalidate()
        ImageBox2.Image = Nothing
        ImageBox3.Image = Nothing
        ImageBox2.BringToFront()
        If _videoSource IsNot Nothing AndAlso _videoSource.IsRunning Then
            Try
                If _videoSource IsNot Nothing AndAlso _videoSource.IsRunning Then
                    ' Ngừng nhận khung hình mới ngay lập tức
                    RemoveHandler _videoSource.NewFrame, AddressOf VideoSource_NewFrame

                    ' Ra lệnh dừng camera
                    _videoSource.SignalToStop()

                    ' Chờ một chút để camera kịp đóng hẳn (tránh lỗi giải phóng bộ nhớ đột ngột)
                    _videoSource.WaitForStop()
                    _videoSource = Nothing
                    ImageBox1.Image = Nothing
                    btnStart.Text = "Lấy từ Camera | OFF"
                End If
            Catch ex As Exception
                ' Bỏ qua lỗi nếu có khi đóng
            End Try
        Else
            Try
                ' 1. Dừng và giải phóng camera cũ nếu đang chạy
                If _videoSource IsNot Nothing AndAlso _videoSource.IsRunning Then
                    RemoveHandler _videoSource.NewFrame, AddressOf VideoSource_NewFrame
                    _videoSource.SignalToStop()
                    _videoSource = Nothing
                End If

                ' 2. Kiểm tra xem người dùng đã chọn camera chưa
                If cbCameras.SelectedIndex = -1 Then
                    MessageBox.Show("Vui lòng chọn một camera!")
                    Return
                End If

                ' 3. Kết nối tới Camera mới
                btnStart.Text = "Lấy từ Camera | ON"
                _videoSource = New VideoCaptureDevice(_videoDevices(cbCameras.SelectedIndex).MonikerString)
                AddHandler _videoSource.NewFrame, AddressOf VideoSource_NewFrame
                ImageBox1.Zoom = 0
                _videoSource.Start()
                ImageBox1.Zoom += 35
            Catch ex As Exception
                btnStart.Text = "Lấy từ Camera | OFF"
                MessageBox.Show("Lỗi: " & ex.Message)
            End Try
        End If
    End Sub




    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If _videoSource IsNot Nothing AndAlso _videoSource.IsRunning Then
                ' Ngừng nhận khung hình mới ngay lập tức
                RemoveHandler _videoSource.NewFrame, AddressOf VideoSource_NewFrame

                ' Ra lệnh dừng camera
                _videoSource.SignalToStop()

                ' Chờ một chút để camera kịp đóng hẳn (tránh lỗi giải phóng bộ nhớ đột ngột)
                _videoSource.WaitForStop()
                _videoSource = Nothing
            End If
        Catch ex As Exception
            ' Bỏ qua lỗi nếu có khi đóng
        End Try
    End Sub




    Private Sub cbCameras_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbCameras.SelectedIndexChanged
        Try
            If _videoSource IsNot Nothing AndAlso _videoSource.IsRunning Then
                ' Ngừng nhận khung hình mới ngay lập tức
                RemoveHandler _videoSource.NewFrame, AddressOf VideoSource_NewFrame

                ' Ra lệnh dừng camera
                _videoSource.SignalToStop()

                ' Chờ một chút để camera kịp đóng hẳn (tránh lỗi giải phóng bộ nhớ đột ngột)
                _videoSource.WaitForStop()
                _videoSource = Nothing
                ImageBox1.Image = Nothing
                btnStart.Text = "Lấy hình từ Camera | OFF"
            End If
        Catch ex As Exception
            ' Bỏ qua lỗi nếu có khi đóng
        End Try

    End Sub

    Private Sub cbCameras_Click(sender As Object, e As EventArgs) Handles cbCameras.Click
        ' 1. Tìm tất cả Camera hiện có trên máy
        _videoDevices = New FilterInfoCollection(FilterCategory.VideoInputDevice)

        If _videoDevices.Count > 0 Then
            ' Đưa tên các Camera vào ComboBox
            cbCameras.Items.Clear()

            For Each device As FilterInfo In _videoDevices
                cbCameras.Items.Add(device.Name)
            Next
            cbCameras.SelectedIndex = 0 ' Chọn cái đầu tiên mặc định
        Else
            MessageBox.Show("Không tìm thấy Camera nào!")
        End If

    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim zoomW As Double, zoomH As Double, fitZoom2 As Integer, fitZoom3 As Integer
        ' Kiểm tra đầu vào
        If ImageBox1.SelectionRegion = RectangleF.Empty OrElse ImageBox1.Image Is Nothing Then
            MessageBox.Show("Vui lòng quét chọn vùng trên hình!")
            Exit Sub
        End If

        Dim sourceRect As Rectangle = Rectangle.Round(ImageBox1.SelectionRegion)
        If sourceRect.Width <= 0 Or sourceRect.Height <= 0 Then Exit Sub

        ' --- XỬ LÝ IMAGEBOX 2 (KHÔNG MÀU - RỘNG 400) ---
        Dim w2 As Integer = 400
        Dim h2 As Integer = CInt((sourceRect.Height / sourceRect.Width) * w2)
        Dim bmp2 As New Bitmap(w2, h2)

        Using g As Graphics = Graphics.FromImage(bmp2)
            SetHighQuality(g)
            ' Vẽ ảnh xám bằng cách dùng ColorMatrix
            Dim grayMatrix As New Drawing.Imaging.ColorMatrix(New Single()() {
            New Single() {0.3, 0.3, 0.3, 0, 0},
            New Single() {0.59, 0.59, 0.59, 0, 0},
            New Single() {0.11, 0.11, 0.11, 0, 0},
            New Single() {0, 0, 0, 1, 0},
            New Single() {0, 0, 0, 0, 1}
        })
            Dim attributes As New Drawing.Imaging.ImageAttributes()
            attributes.SetColorMatrix(grayMatrix)

            g.DrawImage(ImageBox1.Image, New Rectangle(0, 0, w2, h2),
                   sourceRect.X, sourceRect.Y, sourceRect.Width, sourceRect.Height,
                   GraphicsUnit.Pixel, attributes)
        End Using

        ' Hiển thị Box 2
        If ImageBox2.Image IsNot Nothing Then ImageBox2.Image.Dispose()

        ' Tính Zoom vừa khung hình
        zoomW = ImageBox2.ClientSize.Width / bmp2.Width
        zoomH = ImageBox2.ClientSize.Height / bmp2.Height
        fitZoom2 = CInt(Math.Floor(Math.Min(zoomW, zoomH) * 100))
        ImageBox2.Zoom = 0
        ImageBox2.Image = bmp2
        ImageBox2.Zoom = fitZoom2
        ImageBox2.AutoCenter = True
        ' --- XỬ LÝ IMAGEBOX 3 (CÓ MÀU - RỘNG 800) ---
        Dim w3 As Integer = 800
        Dim h3 As Integer = CInt((sourceRect.Height / sourceRect.Width) * w3)
        Dim bmp3 As New Bitmap(w3, h3)

        Using g As Graphics = Graphics.FromImage(bmp3)
            SetHighQuality(g)
            g.DrawImage(ImageBox1.Image, New Rectangle(0, 0, w3, h3), sourceRect, GraphicsUnit.Pixel)
        End Using

        ' Hiển thị Box 3
        If ImageBox3.Image IsNot Nothing Then ImageBox3.Image.Dispose()
        zoomW = ImageBox3.ClientSize.Width / bmp3.Width
        zoomH = ImageBox3.ClientSize.Height / bmp3.Height
        fitZoom3 = CInt(Math.Floor(Math.Min(zoomW, zoomH) * 100))
        ImageBox3.Zoom = 0
        ImageBox3.Image = bmp3
        ImageBox3.Zoom = fitZoom3
        ImageBox3.AutoCenter = True
    End Sub

    ' Hàm thiết lập chất lượng vẽ
    Private Sub SetHighQuality(g As Graphics)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
    End Sub

    Private Sub Set_Text_Name()
        ' Khởi tạo bộ sinh số ngẫu nhiên
        Dim rand As New Random()
        Dim ketQua As String = "IDC_"

        ' Chạy vòng lặp 10 lần để lấy 10 chữ số
        For i As Integer = 1 To 10
            ketQua &= rand.Next(0, 10).ToString()
        Next

        ' Gán giá trị vào Textbox
        TextBox1.Text = ketQua


    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        If Me.Controls.GetChildIndex(Me.ImageBox2) = 0 Then
            Me.ImageBox3.BringToFront()
        Else
            Me.ImageBox2.BringToFront()

        End If
    End Sub



    Private Sub Button_Save_Click(sender As Object, e As EventArgs) Handles Button_Save.Click
        If ImageBox2.Image Is Nothing Then
            MessageBox.Show("Không có hình nào được chọn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Try
            ' 1. Lấy tên file từ TextBox
            Dim fileName As String = TextBox1.Text.Trim()

            ' Kiểm tra nếu TextBox trống
            If String.IsNullOrEmpty(fileName) Then
                MessageBox.Show("Vui lòng nhập tên file trước khi lưu!")
                Return
            End If

            ' Đảm bảo tên file có đuôi .jpg nếu người dùng chưa nhập
            If Not fileName.ToLower().EndsWith(".jpg") Then
                fileName &= ".jpg"
            End If

            ' 2. Đường dẫn các thư mục
            Dim pathUpload As String = "C:\PICTURE_IDC\IDC_UPLOAD"
            Dim pathPrint As String = "C:\PICTURE_IDC\IDC_PRINT"

            ' 3. Kiểm tra và tự động tạo thư mục nếu chưa có
            If Not Directory.Exists(pathUpload) Then Directory.CreateDirectory(pathUpload)
            If Not Directory.Exists(pathPrint) Then Directory.CreateDirectory(pathPrint)

            ' 4. Thực hiện lưu hình ảnh
            ' Lưu ImageBox2 vào IDC_UPLOAD
            If ImageBox2.Image IsNot Nothing Then
                Dim fullPathUpload As String = Path.Combine(pathUpload, fileName)
                'ImageBox2.Image.Save(fullPathUpload, ImageFormat.Jpeg)
                SaveImageWithWhiteBackground(ImageBox2.Image, Path.Combine(pathUpload, fileName))

            End If

            ' Lưu ImageBox3 vào IDC_PRINT
            If ImageBox3.Image IsNot Nothing Then
                Dim fullPathPrint As String = Path.Combine(pathPrint, fileName)
                'ImageBox3.Image.Save(fullPathPrint, ImageFormat.Jpeg)
                SaveImageWithWhiteBackground(ImageBox3.Image, Path.Combine(pathPrint, fileName))

            End If

            MessageBox.Show("Đã lưu hình thành công vào cả 2 thư mục C:\PICTURE_IDC\ ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            renew_img()

        Catch ex As Exception
            MessageBox.Show("Có lỗi xảy ra: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SaveImageWithWhiteBackground(img As Image, fullPath As String)
        ' Tạo Bitmap mới cùng kích thước ảnh gốc
        Using bmp As New Bitmap(img.Width, img.Height)
            Using g As Graphics = Graphics.FromImage(bmp)
                ' TÔ NỀN TRẮNG (Giải quyết vấn đề nền đen)
                g.Clear(Color.White)

                ' Vẽ ảnh đè lên nền trắng
                g.DrawImage(img, New Rectangle(0, 0, img.Width, img.Height))
            End Using

            ' Lưu file định dạng Jpeg
            bmp.Save(fullPath, ImageFormat.Jpeg)
        End Using
    End Sub
    Private Sub renew_img()
        ImageBox1.SelectionRegion = Rectangle.Empty

        ' Cập nhật lại hiển thị để mất khung màu đỏ ngay lập tức
        ImageBox1.Invalidate()
        '
        'If _videoSource IsNot Nothing AndAlso _videoSource.IsRunning Then
        '    ImageBox1.Zoom = 0
        '    ImageBox1.Zoom += 35
        'Else

        'End If
        ImageBox1.Image = Nothing
        ImageBox2.Image = Nothing
        ImageBox3.Image = Nothing
        Set_Text_Name()
    End Sub

#Region "Phong To - thu nho"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnLeft.Click
        ImageBox1.Zoom += 1
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ImageBox1.Zoom -= 1
    End Sub

#End Region


#Region "Di chinh khung chon hinh"

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        SelectionRegion = ImageBox1.SelectionRegion
        SelectionRegion.X -= 5 ' Di chuyển sang trái 5 pixel
        UpdateSelection()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        SelectionRegion = ImageBox1.SelectionRegion
        SelectionRegion.Y += 5
        UpdateSelection()

    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        SelectionRegion = ImageBox1.SelectionRegion
        SelectionRegion.X += 5 ' Di chuyển sang trái 5 pixel
        UpdateSelection()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        SelectionRegion = ImageBox1.SelectionRegion
        SelectionRegion.Width += 5

        UpdateSelection()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        SelectionRegion = ImageBox1.SelectionRegion
        SelectionRegion.Width -= 5

        UpdateSelection()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        SelectionRegion = ImageBox1.SelectionRegion
        SelectionRegion.Height += 5
        UpdateSelection()
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        SelectionRegion = ImageBox1.SelectionRegion
        SelectionRegion.Height -= 5
        UpdateSelection()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        ImageBox1.SelectionRegion = Rectangle.Empty
        ' Cập nhật lại hiển thị để mất khung màu đỏ ngay lập tức
        ImageBox1.Invalidate()

    End Sub

#End Region

End Class
