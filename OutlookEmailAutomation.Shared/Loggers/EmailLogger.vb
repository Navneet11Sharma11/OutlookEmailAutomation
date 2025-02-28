Imports System.IO
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Models

Namespace Loggers
    Public Class EmailLogger
        Implements IEmailLogger

        Public Sub LogError(outlookSettings As OutlookSettings, ex As System.Exception) Implements IEmailLogger.LogError
            If Not outlookSettings.IsLoggingEnabled Then Return

            Dim logDirectory As String = outlookSettings.LogDirectory

            If Not Directory.Exists(logDirectory) Then
                Directory.CreateDirectory(logDirectory)
            End If

            Dim logFilePath As String = Path.Combine(logDirectory, "error.log")

            File.AppendAllText(logFilePath, DateTime.Now.ToString() & ": " & ex.Message & Environment.NewLine)

            Console.WriteLine("Error: " & ex.Message)
        End Sub

        Public Sub LogInfo(outlookSettings As OutlookSettings, message As String) Implements IEmailLogger.LogInfo
            If Not outlookSettings.IsLoggingEnabled Then Return

            Dim logDirectory As String = outlookSettings.LogDirectory

            If Not Directory.Exists(logDirectory) Then
                Directory.CreateDirectory(logDirectory)
            End If

            Dim logFilePath As String = Path.Combine(logDirectory, "info.log")

            File.AppendAllText(logFilePath, DateTime.Now.ToString() & ": " & message & Environment.NewLine)

            Console.WriteLine("Info: " & message)
        End Sub
    End Class
End Namespace