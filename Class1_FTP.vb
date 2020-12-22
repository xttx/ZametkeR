Imports System.Net

Public Class Class1_FTP
    Public host As String
    Public host_arr() As String
    Public last_err As String
    Public Event GetDirectoryDone(res As Boolean, err As String)

    Dim fileList As New List(Of String)

    Function GetFileList() As List(Of String)
        last_err = ""
        If Not IsHostStringValid() Then last_err = "Host configuration is not correct." : Return Nothing

        fileList.Clear()

        Using ftp = New FluentFTP.FtpClient
            ftp.Host = host_arr(0)
            ftp.Credentials = New NetworkCredential(host_arr(1), host_arr(2))
            If Not ftp.DirectoryExists(host_arr(3)) Then last_err = "Requested directory does not exist on the host." : Return Nothing
            GetFileListRecur(ftp, host_arr(3))
        End Using

        Return fileList
    End Function
    Sub GetFileListRecur(ftp As FluentFTP.FtpClient, path As String)
        'Dim l As FluentFTP.FtpListItem() = Await ftp.GetListingAsync(path)
        Dim l As FluentFTP.FtpListItem() = ftp.GetListing(path)
        For Each item In l
            Select Case item.Type
                Case FluentFTP.FtpFileSystemObjectType.Directory
                    GetFileListRecur(ftp, item.FullName)
                Case FluentFTP.FtpFileSystemObjectType.File
                    If item.FullName.ToUpper.EndsWith(".ZAM") Then
                        Dim i = item.FullName
                        If i.ToUpper.StartsWith(host_arr(3).ToUpper) Then i = i.Substring(host_arr(3).Length)
                        fileList.Add(i)
                    End If
            End Select
        Next
    End Sub

    Async Function UploadZam(localPaths As String, remotePath As String) As Task(Of Boolean)
        last_err = ""
        If Not IsHostStringValid() Then last_err = "Host configuration is not correct." : Return False

        Dim loc_arr = localPaths.Split(";"c)

        remotePath = remotePath.Replace("\", "/")
        If Not remotePath.StartsWith("/") Then remotePath = "/" + remotePath
        remotePath = host_arr(3) + remotePath
        remotePath = remotePath.Substring(0, remotePath.LastIndexOf("/"))

        Dim res As Integer = 0
        Using ftp = New FluentFTP.FtpClient
            ftp.Host = host_arr(0)
            ftp.Credentials = New NetworkCredential(host_arr(1), host_arr(2))
            'ftp.UploadFiles(fi, remotePath)
            res = Await ftp.UploadFilesAsync(loc_arr, remotePath)
        End Using

        Return True
    End Function

    Async Function DownloadZam(localPath As String, remotePath As String) As Task(Of Boolean)
        last_err = ""
        If Not IsHostStringValid() Then last_err = "Host configuration is not correct." : Return False

        remotePath = remotePath.Replace("\", "/")
        If Not remotePath.StartsWith("/") Then remotePath = "/" + remotePath
        remotePath = host_arr(3) + remotePath

        Dim res As Boolean = False
        Using ftp = New FluentFTP.FtpClient
            ftp.Host = host_arr(0)
            ftp.Credentials = New NetworkCredential(host_arr(1), host_arr(2))
            If Not Await ftp.FileExistsAsync(remotePath) Then last_err = "Remote file '" + remotePath + "' does not exist." : Return False
            res = Await ftp.DownloadFileAsync(localPath, remotePath)
        End Using
        Return res
    End Function

    Function IsHostStringValid() As Boolean
        Dim path As String = "/"
        Dim r = New System.Text.RegularExpressions.Regex("([A-Za-z0-9\-\.]+):([A-Za-z0-9\-\.]+)@([A-Za-z0-9/\-\.]+)$").Match(host)

        If r.Success Then
            Dim h = r.Groups(3).Value
            If host.Contains("/") Then
                h = h.Substring(0, h.IndexOf("/"))
                path = host.Substring(host.IndexOf("/"))
            End If
            host_arr = {h, r.Groups(1).Value, r.Groups(2).Value, path}
        Else
            host_arr = Nothing
        End If
        Return r.Success
    End Function
End Class
