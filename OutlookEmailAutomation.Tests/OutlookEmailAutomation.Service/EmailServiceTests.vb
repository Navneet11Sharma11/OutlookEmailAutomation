Imports Xunit
Imports Moq
Imports Microsoft.Office.Interop.Outlook
Imports OutlookEmailAutomation.Shared
Imports OutlookEmailAutomation.Business
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Models

Public Class EmailServiceTests

    Private outlookSettings As New OutlookSettings() With {
        .EmailAddress = "test@example.com",
        .SourceFolder = "Inbox",
        .DestinationFolder = "Processed",
        .OutputDirectory = "C:\Output",
        .LogDirectory = "C:\Logs",
        .IsLoggingEnabled = True
    }

    <Fact>
    Public Sub ConnectAndProcessEmails_ValidConnection_ProcessesEmails()
        ' Arrange
        Dim mockApplication As New Mock(Of Application)()
        Dim mockNamespace As New Mock(Of [NameSpace])()
        Dim mockLogger As New Mock(Of IEmailLogger)()
        Dim mockFolder As New Mock(Of MAPIFolder)()
        Dim mockItems As New Mock(Of Items)()
        Dim mockMailItem As New Mock(Of MailItem)()

        mockApplication.Setup(Function(app) app.GetNamespace("MAPI")).Returns(mockNamespace.Object)
        mockNamespace.Setup(Sub(ns) ns.Logon(It.IsAny(Of String), It.IsAny(Of String), It.IsAny(Of Object), It.IsAny(Of Object)))
        mockNamespace.Setup(Function(ns) ns.Folders(outlookSettings.EmailAddress)).Returns(mockFolder.Object)
        mockFolder.Setup(Function(f) f.Folders(outlookSettings.SourceFolder)).Returns(mockFolder.Object)
        mockFolder.Setup(Function(f) f.Items).Returns(mockItems.Object)
        mockItems.Setup(Function(i) i.Count).Returns(1)
        mockItems.Setup(Function(i) i.Item(1)).Returns(mockMailItem.Object)

        Dim emailService As New EmailService(outlookSettings, mockLogger.Object)

        ' Inject mock Application into EmailService
        emailService.OutlookApplication = mockApplication.Object

        ' Act
        emailService.ConnectAndProcessEmails()

        ' Assert
        mockLogger.Verify(Sub(logger) logger.LogInfo(outlookSettings, "Connected to Outlook"), Times.Once)
        mockNamespace.Verify(Sub(ns) ns.Logon(It.IsAny(Of String), It.IsAny(Of String), It.IsAny(Of Object), It.IsAny(Of Object)), Times.Once)
        mockNamespace.Verify(Function(ns) ns.Folders(outlookSettings.EmailAddress), Times.AtLeastOnce)
        mockFolder.Verify(Function(f) f.Folders(outlookSettings.SourceFolder), Times.AtLeastOnce)
        mockFolder.Verify(Function(f) f.Items, Times.AtLeastOnce)
        mockItems.Verify(Function(i) i.Item(1), Times.Once)
    End Sub

    <Fact>
    Public Sub ProcessEmail_ValidEmail_SavesEmailContentToFile()
        ' Arrange
        Dim mockLogger As New Mock(Of IEmailLogger)()
        Dim mockEmail As New Mock(Of MailItem)()
        Dim mockFileSystem As New Mock(Of IFileSystem)()
        Dim expectedContent As String = "Test Email Body"

        mockEmail.Setup(Function(e) e.Body).Returns(expectedContent)
        mockFileSystem.Setup(Sub(fs) fs.WriteAllText(It.IsAny(Of String), It.IsAny(Of String)))

        Dim emailService As New EmailService(outlookSettings, mockLogger.Object, mockFileSystem.Object)
        Dim fileName As String = "Email_" & Guid.NewGuid().ToString() & ".txt"

        ' Act
        emailService.ProcessEmail(mockEmail.Object)

        ' Assert
        mockFileSystem.Verify(Sub(fs) fs.WriteAllText(It.IsAny(Of String), expectedContent), Times.Once)
        mockLogger.Verify(Sub(logger) logger.LogInfo(outlookSettings, It.IsAny(Of String)), Times.Once)
    End Sub

    <Fact>
    Public Sub ReadEmailsFromFolder_ValidFolder_ProcessesAllEmails()
        ' Arrange
        Dim mockNamespace As New Mock(Of [NameSpace])()
        Dim mockFolder As New Mock(Of MAPIFolder)()
        Dim mockItems As New Mock(Of Items)()
        Dim mockEmail As New Mock(Of MailItem)()
        Dim mockLogger As New Mock(Of IEmailLogger)()

        mockNamespace.Setup(Function(ns) ns.Folders(outlookSettings.EmailAddress).Folders(outlookSettings.SourceFolder)).Returns(mockFolder.Object)
        mockFolder.Setup(Function(f) f.Items).Returns(mockItems.Object)
        mockItems.Setup(Function(i) i.Count).Returns(1)
        mockItems.Setup(Function(i) i.Item(1)).Returns(mockEmail.Object)

        Dim emailService As New EmailService(outlookSettings, mockLogger.Object)

        ' Act
        emailService.ReadEmailsFromFolder(mockNamespace.Object)

        ' Assert
        mockLogger.Verify(Sub(logger) logger.LogInfo(outlookSettings, It.IsAny(Of String)), Times.AtLeastOnce)
        mockEmail.Verify(Function(e) e.Body, Times.Once)
    End Sub

    <Fact>
    Public Sub MoveEmailToFolder_ValidEmail_MovesEmailToDestinationFolder()
        ' Arrange
        Dim mockEmail As New Mock(Of MailItem)()
        Dim mockNamespace As New Mock(Of [NameSpace])()
        Dim mockDoneFolder As New Mock(Of MAPIFolder)()
        Dim mockLogger As New Mock(Of IEmailLogger)()

        mockNamespace.Setup(Function(ns) ns.Folders(outlookSettings.EmailAddress).Folders(outlookSettings.DestinationFolder)).Returns(mockDoneFolder.Object)

        Dim emailService As New EmailService(outlookSettings, mockLogger.Object)

        ' Act
        emailService.MoveEmailToFolder(mockEmail.Object, mockNamespace.Object, outlookSettings.DestinationFolder)

        ' Assert
        mockEmail.Verify(Sub(email) email.Move(mockDoneFolder.Object), Times.Once)
        mockLogger.Verify(Sub(logger) logger.LogInfo(outlookSettings, $"Email moved to folder: {outlookSettings.DestinationFolder}"), Times.Once)
    End Sub

End Class
