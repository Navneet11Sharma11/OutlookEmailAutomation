Imports System.IO
Imports OutlookEmailAutomation.Shared

Public Class FileSystemWrapper
    Implements IFileSystem

    Public Function DirectoryExists(path As String) As Boolean Implements IFileSystem.DirectoryExists
        Return Directory.Exists(path)
    End Function

    Public Sub CreateDirectory(path As String) Implements IFileSystem.CreateDirectory
        Directory.CreateDirectory(path)
    End Sub

    Public Sub WriteAllText(path As String, contents As String) Implements IFileSystem.WriteAllText
        File.WriteAllText(path, contents)
    End Sub
End Class
