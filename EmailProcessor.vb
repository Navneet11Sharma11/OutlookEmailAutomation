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

End Class
