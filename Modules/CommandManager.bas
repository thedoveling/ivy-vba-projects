' CommandManager.bas
Option Explicit

' Initialize a new command with the specified SQL text
' sqlText: SQL statement to set as CommandText
' connection: Active database connection
Public Function CreateCommand(sqlText As String, connection As ADODB.connection) As ADODB.Command
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = connection
    cmd.CommandText = sqlText
    cmd.CommandType = adCmdText
    Set CreateCommand = cmd
End Function

' Add a parameter to the specified command
' cmd: Command object to add the parameter to
' name: Parameter name
' dataType: Data type of the parameter (e.g., adVarChar, adInteger)
' size: Size of the parameter, required for text types
' value: Value to assign to the parameter
Public Sub AddParameter(cmd As ADODB.Command, name As String, dataType As DataTypeEnum, size As Long, value As Variant)
    Dim param As ADODB.Parameter
    Set param = cmd.CreateParameter(name, dataType, adParamInput, size, value)
    cmd.Parameters.Append param
End Sub

' Execute a command that returns a recordset (e.g., SELECT queries)
' cmd: Command object with the query
' Returns: Recordset containing query results
Public Function ExecuteQuery(cmd As ADODB.Command) As ADODB.Recordset
    On Error GoTo QueryError
    Dim rs As ADODB.Recordset
    Set rs = cmd.Execute
    Set ExecuteQuery = rs
    Exit Function

QueryError:
    MsgBox "Error executing query: " & Err.Description, vbCritical
    Set ExecuteQuery = Nothing
End Function

' Execute a command that does not return a recordset (e.g., INSERT, UPDATE, DELETE)
' cmd: Command object with the SQL statement
Public Sub ExecuteNonQuery(cmd As ADODB.Command)
    On Error GoTo NonQueryError
    cmd.Execute
    Exit Sub

NonQueryError:
    MsgBox "Error executing non-query command: " & Err.Description, vbCritical
End Sub
