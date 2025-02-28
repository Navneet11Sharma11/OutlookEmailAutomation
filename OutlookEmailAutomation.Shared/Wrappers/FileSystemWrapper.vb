Imports System.IO

Namespace Wrappers
    Public Class FileSystemWrapper
        Implements IFileSystem

        Public Function DirectoryExists(path As String) As Boolean Implements IFileSystem.DirectoryExists
            Return Directory.Exists(path)
        End Function

        Public Sub CreateDirectory(path As String) Implements IFileSystem.CreateDirectory
            Directory.CreateDirectory(path)
        End Sub

        Public Function FileExists(path As String) As Boolean Implements IFileSystem.FileExists
            Return File.Exists(path)
        End Function

        Public Function ReadFile(path As String) As String Implements IFileSystem.ReadFile
            Return File.ReadAllText(path)
        End Function

        Public Sub WriteAllText(path As String, contents As String) Implements IFileSystem.WriteAllText
            File.WriteAllText(path, contents)
        End Sub
    End Class
End Namespace
