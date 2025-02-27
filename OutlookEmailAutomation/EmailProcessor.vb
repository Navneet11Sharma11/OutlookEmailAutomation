Imports System.IO
Imports Microsoft.Office.Interop.Outlook
Imports Newtonsoft.Json

Public Class EmailProcessor
    Public outlookSettings As OutlookSettings

    Public Sub ReadConfiguration()
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), "settings.json")

        If File.Exists(configFilePath) Then
            Dim json As String = File.ReadAllText(configFilePath)
            Dim root As ConfigurationRoot = JsonConvert.DeserializeObject(Of ConfigurationRoot)(json)
            outlookSettings = root.OutlookEmail
            Console.WriteLine("Configuration loaded successfully.")
        Else
            Throw New FileNotFoundException("Configuration file not found!")
        End If
    End Sub

    Public Sub ProcessEmail(email As MailItem)
        Dim emailContent As String = email.Body
        Dim fileName As String = "Email_" & Guid.NewGuid().ToString() & ".txt"
        File.WriteAllText(Path.Combine(outlookSettings.OutputDirectory, fileName), emailContent)
        Console.WriteLine($"Email saved to file: {fileName}")
    End Sub

    Public Sub ReadEmailsFromFolder(ns As [NameSpace])
        Console.WriteLine($"Reading emails from folder: {outlookSettings.SourceFolder}")
        Logger.LogInfo(outlookSettings, $"Reading emails from folder: {outlookSettings.SourceFolder}")
        Dim inboxFolder As MAPIFolder = ns.Folders(outlookSettings.EmailAddress).Folders(outlookSettings.SourceFolder)
        Dim items As Items = inboxFolder.Items
        For i As Integer = items.Count To 1 Step -1
            Dim item As Object = items(i)
            If TypeOf item Is MailItem Then
                Dim email As MailItem = CType(item, MailItem)
                Console.WriteLine($"Processing email: {email.Subject}")
                Logger.LogInfo(outlookSettings, $"Processing email: {email.Subject}")
                ProcessEmail(email)
                MoveEmailToFolder(email, ns, outlookSettings.DestinationFolder)
            End If
        Next
        Console.WriteLine("Finished reading emails.")
        Logger.LogInfo(outlookSettings, "Finished reading emails.")
    End Sub

    Public Sub MoveEmailToFolder(email As MailItem, ns As [NameSpace], destinationFolder As String)
        Dim doneFolder As MAPIFolder = ns.Folders(destinationFolder)
        email.Move(doneFolder)
        Console.WriteLine($"Email moved to folder: {destinationFolder}")
        Logger.LogInfo(outlookSettings, $"Email moved to folder: {destinationFolder}")
    End Sub

End Class
