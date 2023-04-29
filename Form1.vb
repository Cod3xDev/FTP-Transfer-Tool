Imports System.IO
Imports System.Net

Public Class Form1

    ' Dim startTime As DateTime = DateTime.Now
    Private ftpServerUri As String = My.Settings.FtpServerUri
    Private localFolderPath As String = My.Settings.localFolderPath
    Private Sub btnSync_Click(sender As Object, e As EventArgs) Handles btnSync.Click
        ' Dim localFolderPath As String = "C:\Users\Cod3xDev\AppData\Local\bcml\merged_nx" ' Replace with the path to your local folder
        ' Dim ftpServerUri As String = "ftp://192.168.50.165:5000/atmosphere/contents" ' Replace with the URI of your FTP server
        Dim ftpUsername As String = "" ' Replace with your FTP username
        Dim ftpPassword As String = "" ' Replace with your FTP password

        ' Get the list of local files and folders
        Dim localFilesAndFolders = New DirectoryInfo(localFolderPath).GetFileSystemInfos()

        ' Create an FTP request to connect to the server
        Dim request As FtpWebRequest = CType(WebRequest.Create(ftpServerUri), FtpWebRequest)
        request.Credentials = New NetworkCredential(ftpUsername, ftpPassword)
        request.Method = WebRequestMethods.Ftp.ListDirectory

        ' Get the list of files and folders on the server
        Dim response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
        Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
        Dim serverFilesAndFolders = New List(Of String)
        Do While Not reader.EndOfStream
            serverFilesAndFolders.Add(reader.ReadLine())
        Loop
        reader.Close()
        response.Close()

        Dim deleteFolders As New List(Of String) From {"01007EF00011E000", "01007EF00011F001"}

        For Each folder In deleteFolders
            Dim folderUri As New Uri($"{ftpServerUri}/{folder}")
            Dim folderRequest As FtpWebRequest = CType(WebRequest.Create(folderUri), FtpWebRequest)
            folderRequest.Credentials = New NetworkCredential(ftpUsername, ftpPassword)
            folderRequest.Method = WebRequestMethods.Ftp.ListDirectory
            Try
                Using folderResponse As FtpWebResponse = CType(folderRequest.GetResponse(), FtpWebResponse)
                    ' Folder exists, so delete it
                    Dim deleteFolderRequest As FtpWebRequest = CType(WebRequest.Create(folderUri), FtpWebRequest)
                    deleteFolderRequest.Credentials = New NetworkCredential(ftpUsername, ftpPassword)
                    deleteFolderRequest.Method = WebRequestMethods.Ftp.RemoveDirectory
                    Using deleteFolderResponse As FtpWebResponse = CType(deleteFolderRequest.GetResponse(), FtpWebResponse)
                        ' Folder deleted successfully
                    End Using
                End Using
            Catch ex As WebException
                Dim response2 As FtpWebResponse = CType(ex.Response, FtpWebResponse)
                If response.StatusCode = FtpStatusCode.ActionNotTakenFileUnavailable Then
                    ' Folder doesn't exist, so do nothing
                Else
                    ' MessageBox.Show($"Error deleting folder {folder}: {ex.Message}")
                End If
            End Try
        Next

        ' Upload any files that are present locally but not on the server
        For Each localFileOrFolder In localFilesAndFolders
            If localFileOrFolder.Attributes.HasFlag(FileAttributes.Directory) Then
                ' Create the directory on the server if it doesn't exist
                Dim directoryRequest As FtpWebRequest = CType(WebRequest.Create($"{ftpServerUri}/{localFileOrFolder.Name}"), FtpWebRequest)
                directoryRequest.Credentials = New NetworkCredential(ftpUsername, ftpPassword)
                directoryRequest.Method = WebRequestMethods.Ftp.MakeDirectory
                Try
                    Dim directoryResponse As FtpWebResponse = CType(directoryRequest.GetResponse(), FtpWebResponse)
                    directoryResponse.Close()
                Catch ex As Exception
                    ' Ignore the error if the directory already exists
                    If Not ex.Message.StartsWith("550") Then
                        ' MessageBox.Show($"Error creating directory {localFileOrFolder.Name}: {ex.Message}")
                    End If
                End Try
                ' Recursively upload the files in the directory
                UploadDirectory($"{localFolderPath}\{localFileOrFolder.Name}", $"{ftpServerUri}/{localFileOrFolder.Name}", ftpUsername, ftpPassword)
            Else
                ' Upload the file if it doesn't exist on the server or if it has been modified locally
                Dim uploadRequest As FtpWebRequest = CType(WebRequest.Create($"{ftpServerUri}/{localFileOrFolder.Name}"), FtpWebRequest)
                uploadRequest.Credentials = New NetworkCredential(ftpUsername, ftpPassword)
                uploadRequest.Method = WebRequestMethods.Ftp.UploadFile
                uploadRequest.UseBinary = True
                uploadRequest.KeepAlive = False
                uploadRequest.Timeout = 60000 ' 1 minute
                Using localStream As Stream = File.OpenRead($"{localFolderPath}\{localFileOrFolder.Name}"),
                  serverStream As Stream = uploadRequest.GetRequestStream()
                    localStream.CopyTo(serverStream)
                    serverStream.Close()
                End Using
            End If
        Next
        ' Dim endTime As DateTime = DateTime.Now
        ' Dim elapsedTime As TimeSpan = endTime - startTime
        MessageBox.Show($"Sync completed.")
    End Sub

    Private Sub UploadDirectory(localDirectoryPath As String, serverDirectoryUri As String, ftpUsername As String, ftpPassword As String)
        ' Get the list of local files and folders
        Dim localFilesAndFolders = New DirectoryInfo(localDirectoryPath).GetFileSystemInfos()

        ' Create the directory on the server if it doesn't exist
        Dim directoryRequest As FtpWebRequest = CType(WebRequest.Create(serverDirectoryUri), FtpWebRequest)
        directoryRequest.Credentials = New NetworkCredential(ftpUsername, ftpPassword)
        directoryRequest.Method = WebRequestMethods.Ftp.MakeDirectory
        Try
            Dim directoryResponse As FtpWebResponse = CType(directoryRequest.GetResponse(), FtpWebResponse)
            directoryResponse.Close()
        Catch ex As Exception
            ' Ignore the error if the directory already exists
            If Not ex.Message.StartsWith("550") Then
                ' MessageBox.Show($"Error creating directory {serverDirectoryUri}: {ex.Message}")
            End If
        End Try

        ' Upload any files that are present locally but not on the server
        For Each localFileOrFolder In localFilesAndFolders
            If localFileOrFolder.Attributes.HasFlag(FileAttributes.Directory) Then
                ' Recursively upload the files in the directory
                UploadDirectory($"{localDirectoryPath}\{localFileOrFolder.Name}", $"{serverDirectoryUri}/{localFileOrFolder.Name}", ftpUsername, ftpPassword)
            Else
                ' Upload the file if it doesn't exist on the server or if it has been modified locally
                Dim uploadRequest As FtpWebRequest = CType(WebRequest.Create($"{serverDirectoryUri}/{localFileOrFolder.Name}"), FtpWebRequest)
                uploadRequest.Credentials = New NetworkCredential(ftpUsername, ftpPassword)
                uploadRequest.Method = WebRequestMethods.Ftp.UploadFile
                uploadRequest.UseBinary = True
                uploadRequest.KeepAlive = False
                uploadRequest.Timeout = 60000 ' 1 minute

                Using localStream As Stream = File.OpenRead($"{localDirectoryPath}\{localFileOrFolder.Name}"),
                  serverStream As Stream = uploadRequest.GetRequestStream()
                    localStream.CopyTo(serverStream)
                    serverStream.Close()
                End Using
            End If
        Next
    End Sub
    ' Code to show popup and save entered value
    Private Sub ToolStripLabel1_Click(sender As Object, e As EventArgs) Handles ToolStripLabel1.Click
        Dim serverUri As String = My.Settings.FtpServerUri
        Dim prompt As String = "Enter the FTP server URL:"
        Dim title As String = "FTP Server URL"
        serverUri = InputBox(prompt, title, serverUri)

        If Not String.IsNullOrEmpty(serverUri) Then
            My.Settings.FtpServerUri = serverUri
            My.Settings.Save()
        End If
    End Sub

    ' Code to check for saved value at application startup
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim serverUri As String = My.Settings.FtpServerUri
        If String.IsNullOrEmpty(serverUri) Then
            ' Show popup to set FTP server URI
            ToolStripLabel1.PerformClick()
        Else
            ' Use saved FTP server URI
            ftpServerUri = serverUri
        End If
        If String.IsNullOrEmpty(localFolderPath) Then
            ' Show popup to set FTP server URI
            ToolStripLabel2.PerformClick()
            Me.Close()
        Else
            ' Use saved FTP server URI
            localFolderPath = localFolderPath
        End If

    End Sub

    Private Sub ToolStripLabel2_Click(sender As Object, e As EventArgs) Handles ToolStripLabel2.Click
        Dim localFolderPath As String = My.Settings.localFolderPath
        Dim prompt As String = "Enter the BCML Merged Export Directory"
        Dim title As String = "BCML Directory"
        localFolderPath = InputBox(prompt, title, localFolderPath)

        If Not String.IsNullOrEmpty(localFolderPath) Then
            My.Settings.localFolderPath = localFolderPath
            My.Settings.Save()
        End If
    End Sub
End Class