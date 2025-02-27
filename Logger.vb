Imports System.IO

Module Logger
    Sub LogError(ex As System.Exception)
        Dim logDirectory As String = Program.outlookSettings.LogDirectory

        Dim logFilePath As String = Path.Combine(logDirectory, "error.log")

        File.AppendAllText(logFilePath, DateTime.Now.ToString() & ": " & ex.Message & Environment.NewLine)

        Console.WriteLine("Error: " & ex.Message)
    End Sub

    Sub LogInfo(message As String)
        Dim logDirectory As String = Program.outlookSettings.LogDirectory

        Dim logFilePath As String = Path.Combine(logDirectory, "info.log")

        File.AppendAllText(logFilePath, DateTime.Now.ToString() & ": " & message & Environment.NewLine)

        Console.WriteLine("Info: " & message)
    End Sub
End Module