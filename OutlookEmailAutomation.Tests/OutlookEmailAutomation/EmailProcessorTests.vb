Imports System.IO
Imports Moq
Imports Newtonsoft.Json
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Models
Imports Xunit
Imports OutlookEmailAutomation.Shared

Public Class EmailProcessorTests
    Private ReadOnly _mockFileSystem As Mock(Of IFileSystem)
    Private ReadOnly _mockEmailService As Mock(Of IEmailService)
    Private ReadOnly _mockEmailLogger As Mock(Of IEmailLogger)
    Private ReadOnly _emailProcessor As EmailProcessor
    Private ReadOnly _outlookSettings As OutlookSettings

    Public Sub New()
        _mockFileSystem = New Mock(Of IFileSystem)()
        _mockEmailService = New Mock(Of IEmailService)()
        _mockEmailLogger = New Mock(Of IEmailLogger)()
        _emailProcessor = New EmailProcessor(_mockFileSystem.Object)
        _outlookSettings = New OutlookSettings With {
            .OutputDirectory = "output",
            .LogDirectory = "logs"
        }
    End Sub

    <Fact>
    Public Sub ReadConfiguration_ConfigurationFileExists_ShouldLoadSettings()
        ' Arrange
        Dim environment As String = "Development"
        System.Environment.SetEnvironmentVariable("ASPNETCORE_ENVIRONMENT", environment)
        Dim configFileName As String = $"settings.{environment}.json"
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), configFileName)
        Dim configJson As String = JsonConvert.SerializeObject(New With {Key .OutlookEmail = _outlookSettings})

        _mockFileSystem.Setup(Function(fs) fs.FileExists(configFilePath)).Returns(True)
        _mockFileSystem.Setup(Function(fs) fs.ReadFile(configFilePath)).Returns(configJson)
        _mockEmailLogger.Setup(Sub(logger) logger.LogInfo(It.IsAny(Of OutlookSettings)(), It.IsAny(Of String)()))

        ' Act
        _emailProcessor.ReadConfiguration()

        ' Assert
        Assert.NotNull(_emailProcessor.outlookSettings)
        Assert.Equal("output", _emailProcessor.outlookSettings.OutputDirectory)
        Assert.Equal("logs", _emailProcessor.outlookSettings.LogDirectory)
    End Sub

    <Fact>
    Public Sub ReadConfiguration_ConfigurationFileDoesNotExist_ShouldThrowFileNotFoundException()
        ' Arrange
        Dim environment As String = "Development"
        System.Environment.SetEnvironmentVariable("ASPNETCORE_ENVIRONMENT", environment)
        Dim configFileName As String = $"settings.{environment}.json"
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), configFileName)

        _mockFileSystem.Setup(Function(fs) fs.FileExists(configFilePath)).Returns(False)

        ' Act & Assert
        Assert.Throws(Of FileNotFoundException)(Sub() _emailProcessor.ReadConfiguration())
    End Sub

    <Fact>
    Public Sub EnsureDirectoriesExist_DirectoriesDoNotExist_ShouldCreateDirectories()
        ' Arrange
        _emailProcessor.outlookSettings = _outlookSettings

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
        _emailProcessor.outlookSettings = _outlookSettings

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
        _emailProcessor.emailService = _mockEmailService.Object

        ' Act
        _emailProcessor.ProcessEmails()

        ' Assert
        _mockEmailService.Verify(Sub(es) es.ConnectAndProcessEmails(), Times.Once)
    End Sub

    <Fact>
    Public Sub ProcessEmails_EmailServiceNotInitialized_ShouldNotCallConnectAndProcessEmails()
        ' Arrange
        _emailProcessor.emailService = Nothing

        ' Act
        _emailProcessor.ProcessEmails()

        ' Assert
        _mockEmailService.Verify(Sub(es) es.ConnectAndProcessEmails(), Times.Never)
    End Sub
End Class
