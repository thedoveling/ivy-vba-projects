' ErrorHandler.bas
Option Explicit

Private Const LOG_FILE As String = "C:\Temp\ErrorLog.txt" ' Change path as required

' Disables events and screen updating to improve performance.
Public Sub DisableEvents()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

' Enables events and screen updating.
Public Sub EnableEvents()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Handles runtime errors with centralized logging.
' @param customMessage - Custom error message (optional)
Public Sub HandleRuntimeError(Optional customMessage As String = "")
    Dim errMsg As String
    errMsg = "Runtime Error " & Err.Number & ": " & Err.Description & " at line " & Erl

    If customMessage <> "" Then
        errMsg = customMessage & vbCrLf & errMsg
    End If

    ' Log and notify user
    Call LogError(errMsg, LOG_FILE)
    MsgBox errMsg, vbCritical
End Sub

' Logs an error message to a specified log file.
' @param errMsg - The error message to log
' @param logFile - The path to the log file
Public Sub LogError(errMsg As String, logFile As String)
    On Error Resume Next ' Avoid secondary errors
    Dim fileNum As Integer
    fileNum = FreeFile
    Open logFile For Append As #fileNum
    Print #fileNum, Now & " - " & errMsg
    Close #fileNum
End Sub
