Imports System.IO
Imports Newtonsoft.Json
Imports OutlookEmailAutomation.Service.Services
Imports OutlookEmailAutomation.Shared
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Loggers
Imports OutlookEmailAutomation.Shared.Models
Imports OutlookEmailAutomation.Shared.Wrappers

Public Class EmailProcessor
    Public Property outlookSettings As OutlookSettings
    Public Property emailService As EmailService
    Private ReadOnly fileSystem As IFileSystem

    Public Sub New(Optional fileSystem As IFileSystem = Nothing)
        Me.fileSystem = If(fileSystem, New FileSystemWrapper())
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
            Dim emailLogger As IEmailLogger = New EmailLogger()
            emailLogger.LogInfo(outlookSettings, $"Configuration loaded successfully for {environment}.")
            emailService = New EmailService(outlookSettings, emailLogger, fileSystem)
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
