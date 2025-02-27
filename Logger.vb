Imports System.IO

Module Logger
    Sub LogError(ex As System.Exception)
        Dim logDirectory As String = Program.outlookSettings.LogDirectory

        Dim logFilePath As String = Path.Combine(logDirectory, "error.log")

        File.AppendAllText(logFilePath, DateTime.Now.ToString() & ": " & ex.Message & Environment.NewLine)

        Console.WriteLine("Error: " & ex.Message)
    End Sub
End Module
