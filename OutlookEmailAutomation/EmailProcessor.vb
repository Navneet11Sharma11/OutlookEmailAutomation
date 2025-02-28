Imports OutlookEmailAutomation.Shared
Imports OutlookEmailAutomation.Shared.Models

Public Class EmailProcessor
    Private ReadOnly fileSystem As IFileSystem
    Private ReadOnly emailService As IEmailService
    Private ReadOnly outlookSettings As OutlookSettings

    Public Sub New(fileSystem As IFileSystem, emailService As IEmailService, outlookSettings As OutlookSettings)
        Me.fileSystem = fileSystem
        Me.emailService = emailService
        Me.outlookSettings = outlookSettings
    End Sub

    Public Sub ProcessEmails()
        If emailService IsNot Nothing Then
            emailService.ConnectAndProcessEmails()
        Else
            Console.WriteLine("Email service is not initialized.")
        End If
    End Sub

    Public Sub EnsureDirectoriesExist()
        If Not fileSystem.DirectoryExists(outlookSettings.OutputDirectory) Then
            Console.WriteLine($"Creating output directory: {outlookSettings.OutputDirectory}")
            fileSystem.CreateDirectory(outlookSettings.OutputDirectory)
        End If

        If Not fileSystem.DirectoryExists(outlookSettings.LogDirectory) Then
            Console.WriteLine($"Creating log directory: {outlookSettings.LogDirectory}")
            fileSystem.CreateDirectory(outlookSettings.LogDirectory)
        End If
    End Sub
End Class
