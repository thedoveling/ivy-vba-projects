' ErrorHandler.bas
Option Explicit

' Disables events and screen updating.
Public Sub DisableEvents()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

' Enables events and screen updating.
Public Sub EnableEvents()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Standard error-handling routine.
' @param errMsg - The error message to display
' @param logFile - The path to the log file (optional)
Public Sub HandleError(errMsg As String, Optional logFile As String = "")
    ' Log the error if a log file is specified
    If logFile <> "" Then
        Call LogError(errMsg, logFile)
    End If
    
    ' Provide feedback to the user
    MsgBox "Error: " & errMsg, vbCritical
End Sub

' Logs an error message to a specified log file.
' @param errMsg - The error message to log
' @param logFile - The path to the log file
Private Sub LogError(errMsg As String, logFile As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    Print #fileNum, Now & " - " & errMsg
    Close #fileNum
End Sub

' Handles runtime errors and exceptions.
' @param logFile - The path to the log file (optional)
Public Sub HandleRuntimeError(Optional logFile As String = "")
    Dim errMsg As String
    errMsg = "Runtime Error " & Err.Number & ": " & Err.Description & " in " & _
             VBA.Application.VBE.ActiveCodePane.CodeModule & " at line " & Erl
    
    ' Log the error if a log file is specified
    If logFile <> "" Then
        Call LogError(errMsg, logFile)
    End If
    
    ' Provide feedback to the user
    MsgBox errMsg, vbCritical
End Sub