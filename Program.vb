Imports System.IO
Imports System.Reflection
Imports Microsoft.Office.Interop.Outlook
Imports Newtonsoft.Json

Module Program
    Public outlookSettings As OutlookSettings

    Sub Main(args As String())
        Console.WriteLine("Starting application...")
        ReadConfiguration()

        If Not Directory.Exists(outlookSettings.OutputDirectory) Then
            Console.WriteLine($"Creating output directory: {outlookSettings.OutputDirectory}")
            Directory.CreateDirectory(outlookSettings.OutputDirectory)
        End If

        If Not Directory.Exists(outlookSettings.LogDirectory) Then
            Console.WriteLine($"Creating log directory: {outlookSettings.LogDirectory}")
            Directory.CreateDirectory(outlookSettings.LogDirectory)
        End If

        Dim outlookApp As New Application()
        Dim ns As [NameSpace] = outlookApp.GetNamespace("MAPI")
        ns.Logon("", "", Missing.Value, Missing.Value)
        Console.WriteLine("Connected to Outlook")
        Logger.LogInfo("Connected to Outlook")

        Try
            ReadEmailsFromFolder(ns)
        Catch ex As System.Exception
            Logger.LogError(ex)
            Console.WriteLine("An error occurred while reading emails.")
        End Try
    End Sub

    Sub ReadConfiguration()
        Console.WriteLine("Reading configuration...")
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), "settings.json")

        If File.Exists(configFilePath) Then
            Dim json As String = File.ReadAllText(configFilePath)
            Dim root As ConfigurationRoot = JsonConvert.DeserializeObject(Of ConfigurationRoot)(json)

            outlookSettings = root.OutlookEmail
            Console.WriteLine("Configuration loaded successfully.")
            Logger.LogInfo("Configuration loaded successfully.")
        Else
            Console.WriteLine("Configuration file not found!")
            Logger.LogError(New FileNotFoundException("Configuration file not found!"))
            Environment.Exit(1)
        End If
    End Sub

    Sub ReadEmailsFromFolder(ns As [NameSpace])
        Console.WriteLine($"Reading emails from folder: {outlookSettings.SourceFolder}")
        Logger.LogInfo($"Reading emails from folder: {outlookSettings.SourceFolder}")
        Dim inboxFolder As MAPIFolder = ns.Folders(outlookSettings.EmailAddress).Folders(outlookSettings.SourceFolder)
        Dim items As Items = inboxFolder.Items
        For i As Integer = items.Count To 1 Step -1
            Dim item As Object = items(i)
            If TypeOf item Is MailItem Then
                Dim email As MailItem = CType(item, MailItem)
                Console.WriteLine($"Processing email: {email.Subject}")
                Logger.LogInfo($"Processing email: {email.Subject}")
                ProcessEmail(email)
                MoveEmailToFolder(email, ns)
            End If
        Next
        Console.WriteLine("Finished reading emails.")
        Logger.LogInfo("Finished reading emails.")
    End Sub

    Sub ProcessEmail(email As MailItem)
        Dim emailContent As String = email.Body
        Dim fileName As String = "Email_" & Guid.NewGuid().ToString() & ".txt"
        File.WriteAllText(Path.Combine(outlookSettings.OutputDirectory, fileName), emailContent)
        Console.WriteLine($"Email saved to file: {fileName}")
        Logger.LogInfo($"Email saved to file: {fileName}")
    End Sub

    Sub MoveEmailToFolder(email As MailItem, ns As [NameSpace])
        Dim doneFolder As MAPIFolder = ns.Folders(outlookSettings.EmailAddress).Folders(outlookSettings.DestinationFolder)
        email.Move(doneFolder)
        Console.WriteLine($"Email moved to folder: {outlookSettings.DestinationFolder}")
        Logger.LogInfo($"Email moved to folder: {outlookSettings.DestinationFolder}")
    End Sub
End Module