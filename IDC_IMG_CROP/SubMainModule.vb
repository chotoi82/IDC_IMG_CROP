Imports System.Reflection
Module SubMainModule
    ' Biến để đảm bảo chỉ đăng ký sự kiện 1 lần
    Private _isInitialized As Boolean = False
    Public Sub InitApp()
        If _isInitialized Then Return
        ' 1. Đăng ký nạp DLL từ Resource
        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf ResolveAssembly
        ' 2. Giải nén Native SQLite (Bắt buộc phải có để System.Data.SQLite chạy được)
        ExtractNativeDLL()
        _isInitialized = True
    End Sub
    Private Function ResolveAssembly(sender As Object, args As ResolveEventArgs) As Assembly
        ' Lấy tên assembly bị thiếu
        Dim dllName As String = New AssemblyName(args.Name).Name
        ' Lưu ý: Kiểm tra kỹ tên Resource trong My.Resources (thường dấu "." đổi thành "_")
        Try
            Select Case dllName
                Case "AForge.dll"
                    Return Assembly.Load(My.Resources.AForge)
                Case "AForge.Video.DirectShow.dll"
                    Return Assembly.Load(My.Resources.AForge_Video_DirectShow)
                Case "AForge.Video.dll"
                    Return Assembly.Load(My.Resources.AForge_Video)
                Case "Cyotek.Windows.Forms.ImageBox.dll"
                    Return Assembly.Load(My.Resources.Cyotek_Windows_Forms_ImageBox)
                    'Case "System.Threading.Tasks.Extensions"
                    '    Return Assembly.Load(My.Resources.System_Threading_Tasks_Extensions)
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Sub ExtractNativeDLL()
        Try
            ' 1. Lấy đường dẫn thư mục gốc của App
            Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
            'Dim is64Bit As Boolean = Environment.Is64BitProcess
            ' 2. Danh sách các file cần xuất: { "Tên_File_Đầu_Ra.dll", Dữ_Liệu_Byte }
            Dim filesToExtract As New Dictionary(Of String, Byte()) From {
                {"Cyotek.Windows.Forms.ImageBox.dll", My.Resources.Cyotek_Windows_Forms_ImageBox},
                {"AForge.Video.DirectShow.dll", My.Resources.AForge_Video_DirectShow},
                {"AForge.Video.dll", My.Resources.AForge_Video}
            }
            ' 1. Vòng lặp để xuất từng file
            For Each fileEntry In filesToExtract
                Dim filePath As String = IO.Path.Combine(baseDir, fileEntry.Key)
                ' Kiểm tra nếu file chưa tồn tại thì mới xuất (hoặc xóa đi ghi đè nếu muốn cập nhật)
                If Not IO.File.Exists(filePath) Then
                    IO.File.WriteAllBytes(filePath, fileEntry.Value)
                End If
            Next
        Catch ex As Exception
            ' Xử lý lỗi nếu không có quyền ghi vào thư mục (ví dụ thư mục Program Files)
        End Try
    End Sub

End Module
