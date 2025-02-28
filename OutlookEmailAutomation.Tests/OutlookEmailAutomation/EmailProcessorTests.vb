Imports Xunit
Imports Moq
Imports OutlookEmailAutomation.Shared.Models
Imports OutlookEmailAutomation.Shared

Public Class EmailProcessorTests
    Private ReadOnly _mockFileSystem As Mock(Of IFileSystem)
    Private ReadOnly _mockEmailService As Mock(Of IEmailService)
    Private ReadOnly _emailProcessor As EmailProcessor
    Private ReadOnly _outlookSettings As OutlookSettings

    Public Sub New()
        _mockFileSystem = New Mock(Of IFileSystem)()
        _mockEmailService = New Mock(Of IEmailService)()
        _outlookSettings = New OutlookSettings With {
            .OutputDirectory = "output",
            .LogDirectory = "logs"
        }

        _emailProcessor = New EmailProcessor(_mockFileSystem.Object, _mockEmailService.Object, _outlookSettings)
    End Sub

    <Fact>
    Public Sub EnsureDirectoriesExist_DirectoriesDoNotExist_ShouldCreateDirectories()
        ' Arrange
        _mockFileSystem.Setup(Function(fs) fs.DirectoryExists(_outlookSettings.OutputDirectory)).Returns(False)
        _mockFileSystem.Setup(Function(fs) fs.DirectoryExists(_outlookSettings.LogDirectory)).Returns(False)

        ' Act
        _emailProcessor.EnsureDirectoriesExist()

        ' Assert
        _mockFileSystem.Verify(Sub(fs) fs.CreateDirectory(_outlookSettings.OutputDirectory), Times.Once)
        _mockFileSystem.Verify(Sub(fs) fs.CreateDirectory(_outlookSettings.LogDirectory), Times.Once)
    End Sub

    <Fact>
    Public Sub EnsureDirectoriesExist_DirectoriesExist_ShouldNotCreateDirectories()
        ' Arrange
        _mockFileSystem.Setup(Function(fs) fs.DirectoryExists(_outlookSettings.OutputDirectory)).Returns(True)
        _mockFileSystem.Setup(Function(fs) fs.DirectoryExists(_outlookSettings.LogDirectory)).Returns(True)

        ' Act
        _emailProcessor.EnsureDirectoriesExist()

        ' Assert
        _mockFileSystem.Verify(Sub(fs) fs.CreateDirectory(_outlookSettings.OutputDirectory), Times.Never)
        _mockFileSystem.Verify(Sub(fs) fs.CreateDirectory(_outlookSettings.LogDirectory), Times.Never)
    End Sub

    <Fact>
    Public Sub ProcessEmails_EmailServiceInitialized_ShouldCallConnectAndProcessEmails()
        ' Arrange
        ' Act
        _emailProcessor.ProcessEmails()

        ' Assert
        _mockEmailService.Verify(Sub(es) es.ConnectAndProcessEmails(), Times.Once)
    End Sub

    <Fact>
    Public Sub ProcessEmails_EmailServiceNotInitialized_ShouldNotCallConnectAndProcessEmails()
        ' Arrange
        Dim emailProcessorWithoutService As New EmailProcessor(_mockFileSystem.Object, Nothing, _outlookSettings)

        ' Act
        emailProcessorWithoutService.ProcessEmails()

        ' Assert
        _mockEmailService.Verify(Sub(es) es.ConnectAndProcessEmails(), Times.Never)
    End Sub
End Class
