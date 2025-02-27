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
            emailProcessor.ReadEmailsFromFolder(ns)
        Catch ex As System.Exception
            Logger.LogError(emailProcessor.outlookSettings, ex)
            Console.WriteLine("An error occurred while reading emails.")
        End Try
    End Sub
End Module
