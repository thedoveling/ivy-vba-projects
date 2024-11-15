' HelperFunctions.bas
Option Explicit

' Disables Excel events and screen updating to prevent recursion and improve performance
Public Sub DisableEventsAndScreenUpdating()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

' Enables Excel events and screen updating after operations are completed
Public Sub EnableEventsAndScreenUpdating()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Centralized error handler for database transactions
' @param dbManager - The instance of DatabaseManager to manage connections
' @param transactionSuccess - Boolean indicating if the transaction should commit or rollback
Public Sub HandleTransaction(ByRef dbManager As DatabaseManager, ByVal transactionSuccess As Boolean)
    On Error Resume Next
    If transactionSuccess Then
        dbManager.GetConnection.CommitTrans
    Else
        dbManager.GetConnection.RollbackTrans
        MsgBox "Transaction failed. No changes have been committed.", vbCritical
    End If
End Sub

' Logs error messages to a specified worksheet or external file (for debugging purposes)
' @param errorMsg - The error message to log
Public Sub LogError(errorMsg As String)
    ' Customize the logging mechanism as per your requirement
    ' Example: Logging to a worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Error Log")
    
    ' Find the next available row
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    ' Log the error with a timestamp
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = errorMsg
End Sub

' Centralized error handler for VBA errors
' @param errorDescription - Description of the error to display to the user
Public Sub DisplayError(errorDescription As String)
    MsgBox "An error occurred: " & errorDescription, vbCritical
    LogError errorDescription
End Sub
