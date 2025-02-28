Imports System.IO
Imports Newtonsoft.Json
Imports OutlookEmailAutomation.Shared
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Models

Public Class EmailProcessor
    Public Property outlookSettings As OutlookSettings
    Public Property emailService As IEmailService
    Private ReadOnly fileSystem As IFileSystem
    Private ReadOnly emailLogger As IEmailLogger

    Public Sub New(fileSystem As IFileSystem, emailLogger As IEmailLogger, emailService As IEmailService)
        Me.fileSystem = fileSystem
        Me.emailLogger = emailLogger
        Me.emailService = emailService
    End Sub

    Public Sub ReadConfiguration()
        Dim environment As String = If(System.Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT"), "Development")
        Console.WriteLine($"Current environment: {environment}")
        Dim configFileName As String = $"settings.{environment}.json"
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), configFileName)

        If fileSystem.FileExists(configFilePath) Then
            Dim json As String = fileSystem.ReadFile(configFilePath)
            Dim root As ConfigurationRoot = JsonConvert.DeserializeObject(Of ConfigurationRoot)(json)
            outlookSettings = root.OutlookEmail
            Console.WriteLine($"Configuration loaded successfully for {environment}.")
            emailLogger.LogInfo(outlookSettings, $"Configuration loaded successfully for {environment}.")
        Else
            Throw New FileNotFoundException($"Configuration file '{configFileName}' not found!")
        End If
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
