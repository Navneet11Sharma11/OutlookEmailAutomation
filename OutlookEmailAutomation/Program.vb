Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Loggers

Module Program
    Sub Main(args As String())
        Dim emailProcessor As New EmailProcessor()
        Console.WriteLine("Starting application...")
        emailProcessor.ReadConfiguration()

        emailProcessor.EnsureDirectoriesExist()

        Try
            emailProcessor.ProcessEmails()
        Catch ex As System.Exception
            Dim emailLogger As IEmailLogger = New EmailLogger()
            emailLogger.LogError(emailProcessor.outlookSettings, ex)
            Console.WriteLine("An error occurred while reading emails.")
        End Try
    End Sub
End Module
