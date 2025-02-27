Imports Xunit
Imports Moq
Imports System.IO
Imports Newtonsoft.Json
Imports Microsoft.Office.Interop.Outlook

Public Class EmailProcessorTests

    <Fact>
    Public Sub ReadConfiguration_FileExists_LoadsConfiguration()
        ' Arrange
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), "settings.json")
        Dim expectedOutputDirectory As String = "C:\Output"
        Dim config As New ConfigurationRoot() With {
            .OutlookEmail = New OutlookSettings() With {
                .OutputDirectory = expectedOutputDirectory,
                .LogDirectory = "C:\Logs"
            }
        }
        File.WriteAllText(configFilePath, JsonConvert.SerializeObject(config))
        Dim emailProcessor As New EmailProcessor()

        ' Act
        emailProcessor.ReadConfiguration()

        ' Assert
        Assert.Equal(expectedOutputDirectory, emailProcessor.outlookSettings.OutputDirectory)

        ' Cleanup
        File.Delete(configFilePath)
    End Sub

    <Fact>
    Public Sub ReadConfiguration_FileNotFound_ThrowsFileNotFoundException()
        ' Arrange
        Dim configFilePath As String = Path.Combine(Directory.GetCurrentDirectory(), "settings.json")

        If File.Exists(configFilePath) Then
            File.Delete(configFilePath)
        End If

        Dim emailProcessor As New EmailProcessor()

        ' Act & Assert
        Dim ex = Assert.Throws(Of FileNotFoundException)(Sub() emailProcessor.ReadConfiguration())
        Assert.Equal("Configuration file not found!", ex.Message)
    End Sub

    <Fact>
    Public Sub ProcessEmail_SavesEmailContentToFile()
        ' Arrange
        Dim mockEmail As New Mock(Of MailItem)()
        Dim expectedContent As String = "Test Email Body"
        mockEmail.Setup(Function(e) e.Body).Returns(expectedContent)
        Dim outputDirectory As String = Path.Combine(Directory.GetCurrentDirectory(), "Output")
        Dim outputFilePath As String = Path.Combine(outputDirectory, "Email_Test.txt")
        Directory.CreateDirectory(outputDirectory)
        Dim emailProcessor As New EmailProcessor() With {
            .outlookSettings = New OutlookSettings() With {
                .OutputDirectory = outputDirectory
            }
        }

        ' Act
        emailProcessor.ProcessEmail(mockEmail.Object)

        ' Assert
        Assert.True(File.Exists(outputFilePath))
        Dim actualContent As String = File.ReadAllText(outputFilePath)
        Assert.Equal(expectedContent, actualContent)

        ' Cleanup
        File.Delete(outputFilePath)
        Directory.Delete(outputDirectory)
    End Sub

End Class
