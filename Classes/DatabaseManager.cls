' DatabaseManager.cls
Option Explicit

Private con As ADODB.Connection

' Opens a database connection using the globally stored credentials.
Public Function OpenConnection() As Boolean
    On Error GoTo ConnectionError
    Dim connectionString As String

    ' Ensure global credentials are set
    If UserID = vbNullString Or Password = vbNullString Then
        MsgBox "Error: User credentials are not set. Please set them first.", vbCritical
        OpenConnection = False
        Exit Function
    End If

    ' Build and open the connection
    connectionString = "Provider=ORAOLEDB.ORACLE;" & _
                       "Data Source=DEVVV;" & _
                       "User Id=" & UserID & ";" & _
                       "Password=" & Password & ";"
    Set con = New ADODB.Connection
    con.ConnectionString = connectionString
    con.Open

    isConnected = True
    OpenConnection = True
    Exit Function

ConnectionError:
    MsgBox "Error opening connection: " & Err.Description, vbCritical
    If Not con Is Nothing Then
        If con.State = adStateOpen Then con.Close
        Set con = Nothing
    End If
    'isConnected = False
    OpenConnection = False
End Function

' Checks if the connection is active.
Public Function IsConnected() As Boolean
    On Error Resume Next
    IsConnected = Not con Is Nothing And con.State = adStateOpen
End Function


    ' Build and open the connection
    connectionString = "Provider=ORAOLEDB.ORACLE;" & _
                       "Data Source=DEVVV;" & _
                       "User Id=" & userID & ";" & _
                       "Password=" & password & ";"
    Set con = New ADODB.Connection
    con.ConnectionString = connectionString
    con.Open

    OpenConnection = True
    Exit Function

ConnectionError:
    MsgBox "Error opening connection: " & Err.Description, vbCritical
    If Not con Is Nothing Then
        If con.State = adStateOpen Then con.Close
        Set con = Nothing
    End If
    OpenConnection = False
End Function

' Clears stored credentials.
Public Sub ClearCredentials()
    userID = vbNullString
    password = vbNullString
End Sub

' Closes the database connection if open.
Public Sub CloseConnection()
    If Not con Is Nothing And con.State = adStateOpen Then
        con.Close
        Set con = Nothing
    End If
End Sub



' Returns the active connection.
Public Function GetConnection() As ADODB.Connection
    If con Is Nothing Or con.State <> adStateOpen Then
        MsgBox "Connection is not open.", vbCritical
        Set GetConnection = Nothing
    Else
        Set GetConnection = con
    End If
End Function

' Begins a transaction on the open connection.
Public Sub BeginTransaction()
    If Not con Is Nothing And con.State = adStateOpen Then
        con.BeginTrans
    End If
End Sub

' Commits the current transaction.
Public Sub CommitTransaction()
    If Not con Is Nothing And con.State = adStateOpen Then
        con.CommitTrans
    End If
End Sub

' Rolls back the current transaction.
Public Sub RollbackTransaction()
    If Not con Is Nothing And con.State = adStateOpen Then
        con.RollbackTrans
    End If
End Sub

' Creates a command with the specified SQL text.
' @param sqlText - SQL statement to set as CommandText
' @param connection - Active database connection
' @return - ADODB.Command object
Public Function CreateCommand(sqlText As String, connection As ADODB.Connection) As ADODB.Command
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = connection
    cmd.CommandText = sqlText
    cmd.CommandType = adCmdText
    Set CreateCommand = cmd
End Function

' Adds a parameter to the specified command.
' @param cmd - Command object to add the parameter to
' @param name - Parameter name
' @param dataType - Data type of the parameter (e.g., adVarChar, adInteger)
' @param size - Size of the parameter, required for text types
' @param value - Value to assign to the parameter
Public Sub AddParameter(cmd As ADODB.Command, name As String, dataType As DataTypeEnum, size As Long, value As Variant)
    Dim param As ADODB.Parameter
    Set param = cmd.CreateParameter(name, dataType, adParamInput, size, value)
    cmd.Parameters.Append param
End Sub

' Executes a command that returns a recordset (e.g., SELECT queries).
' @param cmd - Command object with the query
' @return - ADODB.Recordset containing query results
Public Function ExecuteCommandQuery(cmd As ADODB.Command) As ADODB.Recordset
    On Error GoTo QueryError
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open cmd, , adOpenStatic, adLockReadOnly
    Set ExecuteCommandQuery = rs
    Exit Function

QueryError:
    MsgBox "Error executing query: " & Err.Description, vbCritical
    Set ExecuteCommandQuery = Nothing
End Function

' Executes a command that does not return a recordset (e.g., INSERT, UPDATE, DELETE).
' @param cmd - Command object with the SQL statement
Public Sub ExecuteCommandNonQuery(cmd As ADODB.Command)
    On Error GoTo NonQueryError
    cmd.Execute
    Exit Sub

NonQueryError:
    MsgBox "Error executing non-query command: " & Err.Description, vbCritical
End Sub

' Executes a series of commands within a transaction.
' @param commands - Collection of ADODB.Command objects to execute
' @return - Boolean indicating success or failure
Public Function ExecuteTransaction(commands As Collection) As Boolean
    On Error GoTo TransactionError
    Dim cmd As ADODB.Command
    
    ' Begin transaction
    Call BeginTransaction
    
    ' Execute each command
    For Each cmd In commands
        cmd.Execute
    Next cmd
    
    ' Commit transaction
    Call CommitTransaction
    ExecuteTransaction = True
    Exit Function

TransactionError:
    ' Rollback transaction on error
    Call RollbackTransaction
    MsgBox "Error executing transaction: " & Err.Description, vbCritical
    ExecuteTransaction = False
End Function