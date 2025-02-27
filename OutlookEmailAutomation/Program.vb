Imports System.IO
Imports System.Reflection
Imports Microsoft.Office.Interop.Outlook

Module Program
    Sub Main(args As String())
        Dim emailProcessor As New EmailProcessor()
        Console.WriteLine("Starting application...")
        emailProcessor.ReadConfiguration()

        If Not Directory.Exists(emailProcessor.outlookSettings.OutputDirectory) Then
            Console.WriteLine($"Creating output directory: {emailProcessor.outlookSettings.OutputDirectory}")
            Directory.CreateDirectory(emailProcessor.outlookSettings.OutputDirectory)
        End If

        Dim outlookApp As New Application()
        Dim ns As [NameSpace] = outlookApp.GetNamespace("MAPI")
        ns.Logon("", "", Missing.Value, Missing.Value)
        Console.WriteLine("Connected to Outlook")

        Try
            ReadEmailsFromFolder(ns, emailProcessor)
        Catch ex As System.Exception
            Console.WriteLine("An error occurred while reading emails.")
        End Try
    End Sub

    Sub ReadEmailsFromFolder(ns As [NameSpace], emailProcessor As EmailProcessor)
        Console.WriteLine($"Reading emails from folder: {emailProcessor.outlookSettings.SourceFolder}")
        Dim inboxFolder As MAPIFolder = ns.Folders(emailProcessor.outlookSettings.EmailAddress).Folders(emailProcessor.outlookSettings.SourceFolder)
        Dim items As Items = inboxFolder.Items
        For i As Integer = items.Count To 1 Step -1
            Dim item As Object = items(i)
            If TypeOf item Is MailItem Then
                Dim email As MailItem = CType(item, MailItem)
                Console.WriteLine($"Processing email: {email.Subject}")
                emailProcessor.ProcessEmail(email)
                MoveEmailToFolder(email, ns, emailProcessor.outlookSettings.DestinationFolder)
            End If
        Next
        Console.WriteLine("Finished reading emails.")
    End Sub

    Sub MoveEmailToFolder(email As MailItem, ns As [NameSpace], destinationFolder As String)
        Dim doneFolder As MAPIFolder = ns.Folders(destinationFolder)
        email.Move(doneFolder)
        Console.WriteLine($"Email moved to folder: {destinationFolder}")
    End Sub

End Module
