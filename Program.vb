Imports System.IO
Imports System.Reflection
Imports Microsoft.Office.Interop.Outlook
Imports Newtonsoft.Json

Module Program
    Public outlookSettings As OutlookSettings

    Sub Main(args As String())
        ReadConfiguration()

        If Not Directory.Exists(outlookSettings.OutputDirectory) Then
            Directory.CreateDirectory(outlookSettings.OutputDirectory)
        End If

        If Not Directory.Exists(outlookSettings.LogDirectory) Then
            Directory.CreateDirectory(outlookSettings.LogDirectory)
        End If

        Dim outlookApp As New Application()
        Dim ns As [NameSpace] = outlookApp.GetNamespace("MAPI")
        ns.Logon("", "", Missing.Value, Missing.Value)
        Console.WriteLine("Connected to Outlook")

        Try
            ReadEmailsFromFolder(ns)
        Catch ex As System.Exception
            Logger.LogError(ex)
        End Try
    End Sub

    Sub ReadConfiguration()
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), "settings.json")

        If File.Exists(configFilePath) Then
            Dim json As String = File.ReadAllText(configFilePath)
            Dim root As ConfigurationRoot = JsonConvert.DeserializeObject(Of ConfigurationRoot)(json)

            outlookSettings = root.OutlookEmail
        Else
            Console.WriteLine("Configuration file not found!")
            Environment.Exit(1)
        End If
    End Sub

    Sub ReadEmailsFromFolder(ns As [NameSpace])
        Dim inboxFolder As MAPIFolder = ns.Folders(outlookSettings.EmailAddress).Folders(outlookSettings.SourceFolder)
        For Each item As Object In inboxFolder.Items
            If TypeOf item Is MailItem Then
                Dim email As MailItem = CType(item, MailItem)
                ProcessEmail(email)
                MoveEmailToFolder(email, ns)
            End If
        Next
    End Sub

    Sub ProcessEmail(email As MailItem)
        Dim emailContent As String = email.Body
        Dim fileName As String = "Email_" & Guid.NewGuid().ToString() & ".txt"
        File.WriteAllText(Path.Combine(outlookSettings.OutputDirectory, fileName), emailContent)
    End Sub

    Sub MoveEmailToFolder(email As MailItem, ns As [NameSpace])
        Dim doneFolder As MAPIFolder = ns.Folders(outlookSettings.EmailAddress).Folders(outlookSettings.DestinationFolder)
        email.Move(doneFolder)
    End Sub
End Module
