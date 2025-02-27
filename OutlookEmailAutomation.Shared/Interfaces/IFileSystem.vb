Public Interface IFileSystem
    Function DirectoryExists(path As String) As Boolean
    Sub CreateDirectory(path As String)
    Sub WriteAllText(path As String, contents As String)
End Interface
