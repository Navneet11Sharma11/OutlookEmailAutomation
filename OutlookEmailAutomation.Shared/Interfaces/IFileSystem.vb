Public Interface IFileSystem
    Function DirectoryExists(path As String) As Boolean
    Sub CreateDirectory(path As String)
    Function FileExists(path As String) As Boolean
    Function ReadFile(path As String) As String
    Sub WriteAllText(path As String, contents As String)
End Interface
