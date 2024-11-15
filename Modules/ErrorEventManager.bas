' ErrorEventManager.bas
Option Explicit

' Enables or disables all Excel events.
' enable: Boolean flag to enable (True) or disable (False) events.
Public Sub SetEvents(enable As Boolean)
    Application.EnableEvents = enable
End Sub

' Enables or disables screen updating in Excel.
' enable: Boolean flag to enable (True) or disable (False) screen updating.
Public Sub SetScreenUpdating(enable As Boolean)
    Application.ScreenUpdating = enable
End Sub

' General error handler that logs the error and notifies the user.
' errorMsg: Message to display to the user.
' errorLocation: Description of where the error occurred (e.g., function name).
Public Sub HandleError(errorMsg As String, errorLocation As String)
    ' Log the error with details
    LogError errorMsg, errorLocation
    
    ' Display a message box to the user
    MsgBox "An error occurred in " & errorLocation & ": " & errorMsg, vbCritical
End Sub

' Logs an error message to a designated "Error Log" sheet in the workbook.
' errorMsg: Error message to log.
' errorLocation: Location of the error for context.
Public Sub LogError(errorMsg As String, errorLocation As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Try to reference an "Error Log" sheet, or create it if it doesn't exist.
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Error Log")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "Error Log"
        ws.Range("A1").value = "Timestamp"
        ws.Range("B1").value = "Location"
        ws.Range("C1").value = "Error Message"
    End If
    
    ' Find the last empty row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    ' Log error details
    ws.Cells(lastRow, 1).value = Now()
    ws.Cells(lastRow, 2).value = errorLocation
    ws.Cells(lastRow, 3).value = errorMsg
End Sub

' Wrapper function to disable events and screen updating.
Public Sub DisableEventsAndScreenUpdating()
    SetEvents False
    SetScreenUpdating False
End Sub

' Wrapper function to enable events and screen updating.
Public Sub EnableEventsAndScreenUpdating()
    SetEvents True
    SetScreenUpdating True
End Sub
