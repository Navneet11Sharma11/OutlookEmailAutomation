Imports Microsoft.Office.Interop.Outlook
Imports OutlookEmailAutomation.Shared
Imports OutlookEmailAutomation.Shared.Interfaces
Imports OutlookEmailAutomation.Shared.Models
Imports System.IO
Imports System.Reflection

Public Class EmailService
    Implements IEmailService

    Private ReadOnly outlookSettings As OutlookSettings
    Private ReadOnly emailLogger As IEmailLogger
    Private ReadOnly fileSystem As IFileSystem

    Public Property OutlookApplication As Application

    Public Sub New(settings As OutlookSettings, logger As IEmailLogger, Optional fileSystem As IFileSystem = Nothing)
        Me.outlookSettings = settings
        Me.emailLogger = logger
        Me.fileSystem = If(fileSystem, New FileSystemWrapper())
        Me.OutlookApplication = New Application()
    End Sub

    Public Sub ConnectAndProcessEmails() Implements IEmailService.ConnectAndProcessEmails
        Dim ns As [NameSpace] = OutlookApplication.GetNamespace("MAPI")
        ns.Logon("", "", Missing.Value, Missing.Value)
        emailLogger.LogInfo(outlookSettings, "Connected to Outlook")

        Try
            ReadEmailsFromFolder(ns)
        Catch ex As System.Exception
            emailLogger.LogError(outlookSettings, ex)
        End Try
    End Sub

    Public Sub ProcessEmail(email As MailItem)
        Try
            Dim emailContent As String = email.Body
            Dim fileName As String = "Email_" & Guid.NewGuid().ToString() & ".txt"
            Dim filePath As String = Path.Combine(outlookSettings.OutputDirectory, fileName)
            fileSystem.WriteAllText(filePath, emailContent)
            emailLogger.LogInfo(outlookSettings, $"Email saved to file: {filePath}")
        Catch ex As System.Exception
            emailLogger.LogError(outlookSettings, ex)
        End Try
    End Sub

    Public Sub ReadEmailsFromFolder(ns As [NameSpace])
        Dim inboxFolder As MAPIFolder = ns.Folders(outlookSettings.EmailAddress).Folders(outlookSettings.SourceFolder)
        Dim items As Items = inboxFolder.Items
        For i As Integer = items.Count To 1 Step -1
            Dim item As Object = items(i)
            If TypeOf item Is MailItem Then
                Dim email As MailItem = CType(item, MailItem)
                ProcessEmail(email)
                MoveEmailToFolder(email, ns, outlookSettings.DestinationFolder)
            End If
        Next
    End Sub

    Public Sub MoveEmailToFolder(email As MailItem, ns As [NameSpace], destinationFolder As String)
        Try
            Dim doneFolder As MAPIFolder = ns.Folders(outlookSettings.EmailAddress).Folders(destinationFolder)
            email.Move(doneFolder)
            emailLogger.LogInfo(outlookSettings, $"Email moved to folder: {destinationFolder}")
        Catch ex As System.Exception
            emailLogger.LogError(outlookSettings, ex)
            Throw
        End Try
    End Sub
End Class
