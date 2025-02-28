Imports Microsoft.Extensions.DependencyInjection
Imports OutlookEmailAutomation.Shared
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Wrappers
Imports Newtonsoft.Json
Imports System.IO
Imports OutlookEmailAutomation.Service.Services
Imports OutlookEmailAutomation.Shared.Models
Imports OutlookEmailAutomation.Shared.Loggers

Module Program
    Sub Main(args As String())
        Dim outlookSettings As OutlookSettings = LoadConfiguration()

        Dim serviceProvider As IServiceProvider = New ServiceCollection() _
            .AddSingleton(Of IFileSystem, FileSystemWrapper)() _
            .AddSingleton(Of IEmailLogger, EmailLogger)() _
            .AddSingleton(Of IEmailService, EmailService)() _
            .AddSingleton(outlookSettings) _
            .AddSingleton(Of EmailProcessor)() _
            .BuildServiceProvider()

        Dim emailProcessor As EmailProcessor = serviceProvider.GetService(Of EmailProcessor)()
        emailProcessor.EnsureDirectoriesExist()
        emailProcessor.ProcessEmails()
    End Sub

    Private Function LoadConfiguration() As OutlookSettings
        Dim environment As String = If(System.Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT"), "Development")
        Console.WriteLine($"Current environment: {environment}")
        Dim configFileName As String = $"settings.{environment}.json"
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), configFileName)
        Dim outlookSettings As OutlookSettings = Nothing

        If File.Exists(configFilePath) Then
            Dim json As String = File.ReadAllText(configFilePath)
            Dim root As ConfigurationRoot = JsonConvert.DeserializeObject(Of ConfigurationRoot)(json)
            outlookSettings = root.OutlookEmail
            Console.WriteLine($"Configuration loaded successfully for {environment}.")
        Else
            Throw New FileNotFoundException($"Configuration file '{configFileName}' not found!")
        End If

        Return outlookSettings
    End Function
End Module
