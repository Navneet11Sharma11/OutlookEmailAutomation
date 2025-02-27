Imports OutlookEmailAutomation.Shared.Models

Namespace Interfaces
    Public Interface IEmailLogger
        Sub LogError(outlookSettings As OutlookSettings, ex As System.Exception)
        Sub LogInfo(outlookSettings As OutlookSettings, message As String)
    End Interface
End Namespace

