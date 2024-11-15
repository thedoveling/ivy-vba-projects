' QueryBuilder.cls
Option Explicit

' Dependencies
Private dbManager As DatabaseManager

' Initializes the QueryBuilder with a DatabaseManager instance for handling connections
Public Sub Initialize(manager As DatabaseManager)
    Set dbManager = manager
End Sub

' Creates a new ADODB.Parameter with error handling and default size/type management
' @param name - The name of the parameter
' @param paramType - The ADODB data type of the parameter
' @param size - The size of the parameter (optional)
' @param value - The value to assign to the parameter (optional)
' @return - An ADODB.Parameter ready for use
Public Function CreateParameter(name As String, paramType As DataTypeEnum, Optional size As Long = 0, Optional value As Variant) As ADODB.Parameter
    On Error GoTo ErrorHandler
    Dim param As ADODB.Parameter
    If size > 0 Then
        Set param = dbManager.GetConnection.CreateParameter(name, paramType, adParamInput, size, value)
    Else
        Set param = dbManager.GetConnection.CreateParameter(name, paramType, adParamInput, , value)
    End If
    Set CreateParameter = param
    Exit Function
ErrorHandler:
    MsgBox "Error creating parameter '" & name & "': " & Err.Description, vbCritical
    Set CreateParameter = Nothing
End Function

' Executes a command (INSERT, UPDATE, DELETE) with parameters and manages transaction
' @param sqlQuery - The SQL command string
' @param params - A Variant that can hold an array of ADODB.Parameter objects (optional)
' @return - Boolean indicating success (True) or failure (False)
Public Function ExecuteCommand(sqlQuery As String, Optional params As Variant) As Boolean
    On Error GoTo CommandError
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim transactionStarted As Boolean
    transactionStarted = False

    ' Initialize command properties
    Set cmd.ActiveConnection = dbManager.GetConnection
    cmd.CommandText = sqlQuery
    cmd.CommandType = adCmdText

    ' Append parameters if provided and is an array
    If IsArray(params) Then
        Dim i As Integer
        For i = LBound(params) To UBound(params)
            cmd.Parameters.Append params(i)
        Next i
    End If

    ' Begin transaction
    dbManager.GetConnection.BeginTrans
    transactionStarted = True
    cmd.Execute

    ' Commit transaction only if execute is successful
    dbManager.GetConnection.CommitTrans
    ExecuteCommand = True
    Exit Function

CommandError:
    ' Rollback transaction if an error occurs during execution
    If transactionStarted Then dbManager.GetConnection.RollbackTrans
    MsgBox "Error executing command: " & Err.Description, vbCritical
    ExecuteCommand = False
End Function

' Executes a SELECT query and returns an ADODB.Recordset
' @param sqlQuery - The SQL SELECT query string
' @param params - A Variant that can hold an array of ADODB.Parameter objects (optional)
' @return - An ADODB.Recordset containing the results (or Nothing if failed)
Public Function ExecuteSelect(sqlQuery As String, Optional params As Variant) As ADODB.Recordset
    On Error GoTo SelectError
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Set cmd = New ADODB.Command

    ' Initialize command properties
    cmd.ActiveConnection = dbManager.GetConnection
    cmd.CommandText = sqlQuery
    cmd.CommandType = adCmdText

    ' Append parameters if provided and is an array
    If IsArray(params) Then
        Dim i As Integer
        For i = LBound(params) To UBound(params)
            cmd.Parameters.Append params(i)
        Next i
    End If

    ' Execute the command and return the recordset
    Set rs = cmd.Execute
    Set ExecuteSelect = rs
    Exit Function

SelectError:
    MsgBox "Error executing select query: " & Err.Description, vbCritical
    Set ExecuteSelect = Nothing
End Function

