Imports Xunit
Imports Moq
Imports OutlookEmailAutomation.Shared
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Models
Imports OutlookEmailAutomation.Service.Services
Imports Newtonsoft.Json

Public Class ProgramTests

    Private outlookSettings As New OutlookSettings() With {
        .EmailAddress = "test@example.com",
        .SourceFolder = "Inbox",
        .DestinationFolder = "Processed",
        .OutputDirectory = "C:\Emails",
        .LogDirectory = "C:\Logs",
        .IsLoggingEnabled = True
    }

    <Fact>
    Public Sub Main_ValidConfiguration_ProcessesEmails()
        ' Arrange
        Dim mockFileSystem As New Mock(Of IFileSystem)()
        Dim mockLogger As New Mock(Of IEmailLogger)()

        mockFileSystem.Setup(Function(fs) fs.DirectoryExists(It.IsAny(Of String))).Returns(False)
        mockFileSystem.Setup(Sub(fs) fs.CreateDirectory(It.IsAny(Of String)))
        mockFileSystem.Setup(Function(fs) fs.FileExists(It.IsAny(Of String))).Returns(True)

        Dim configJson As String = JsonConvert.SerializeObject(New With {Key .OutlookEmail = outlookSettings})
        mockFileSystem.Setup(Function(fs) fs.ReadFile(It.IsAny(Of String))).Returns(configJson)

        Dim emailProcessor = New EmailProcessor(mockFileSystem.Object) With {
            .outlookSettings = outlookSettings
        }

        emailProcessor.emailService = New EmailService(outlookSettings, mockLogger.Object, mockFileSystem.Object)

        ' Act
        emailProcessor.ReadConfiguration()
        emailProcessor.EnsureDirectoriesExist()
        emailProcessor.ProcessEmails()

        ' Assert
        mockFileSystem.Verify(Function(fs) fs.DirectoryExists(outlookSettings.OutputDirectory), Times.Once)
        mockFileSystem.Verify(Sub(fs) fs.CreateDirectory(outlookSettings.OutputDirectory), Times.Once)
        mockFileSystem.Verify(Function(fs) fs.DirectoryExists(outlookSettings.LogDirectory), Times.Once)
        mockFileSystem.Verify(Sub(fs) fs.CreateDirectory(outlookSettings.LogDirectory), Times.Once)
    End Sub

End Class
