Imports Microsoft.Extensions.DependencyInjection
Imports OutlookEmailAutomation.Service.Services
Imports OutlookEmailAutomation.Shared
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Loggers
Imports OutlookEmailAutomation.Shared.Wrappers

Module Program
    Sub Main(args As String())
        Dim serviceProvider As IServiceProvider = New ServiceCollection() _
            .AddSingleton(Of IFileSystem, FileSystemWrapper)() _
            .AddSingleton(Of IEmailLogger, EmailLogger)() _
            .AddSingleton(Of IEmailService, EmailService)() _
            .AddSingleton(Of EmailProcessor)() _
            .BuildServiceProvider()

        Dim emailProcessor As EmailProcessor = serviceProvider.GetService(Of EmailProcessor)()
        emailProcessor.ReadConfiguration()
        emailProcessor.EnsureDirectoriesExist()
        emailProcessor.ProcessEmails()
    End Sub
End Module
