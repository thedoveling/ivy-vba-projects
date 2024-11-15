' CommunicationManager.bas
Option Explicit

' Routes communication based on type and initialises necessary configuration
' @param ws - Worksheet containing data for placeholders
' @param configType - Type of communication: "Email" or "Word"
' @param startRow - Starting row for data in the worksheet
Public Sub ExecuteCommunication(ws As Worksheet, configType As String, startRow As Long)
    Dim config As Scripting.Dictionary

    ' Retrieve configurations dynamically based on communication type
    Set config = ConfigManager.GetCommunicationConfig(configType)
    
    Select Case UCase(configType)
        Case "EMAIL"
            EmailManager.CreateEmailDrafts ws, startRow, config
        Case "WORD"
            WordMailMergeManager.PerformMailMerge ws, startRow, config
        Case Else
            MsgBox "Unsupported communication type: " & configType, vbCritical
    End Select
End Sub
