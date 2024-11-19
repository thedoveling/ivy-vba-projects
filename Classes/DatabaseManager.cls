'@summary This class represents a database manager that handles connections, transactions, and query execution.
'@remarks This class uses the ADODB library for database operations.

Option Explicit

Private con As ADODB.Connection

Private transactionInProgress As Boolean

'@summary Sets the credentials for the database connection.
'@param uID The user ID for the database connection.
'@param pwd The password for the database connection.
Public Sub SetCredentials(uID As String, pwd As String)
    userID = uID
    password = pwd
End Sub

'@summary Opens a connection to the database.
'@returns True if the connection is successfully opened, False otherwise.
Public Function OpenConnection() As Boolean
    On Error GoTo ConnectionError

    If Not IsConnected Then
        Dim connectionString As String
        connectionString = "Provider=ORAOLEDB.ORACLE;" & _
                           "Data Source=DEVVV;" & _
                           "User Id=" & userID & ";" & _
                           "Password=" & password & ";"
        Set con = New ADODB.Connection
        con.ConnectionString = connectionString
        con.Open
    End If

    OpenConnection = IsConnected
    Exit Function

ConnectionError:
    MsgBox "Error opening connection: " & Err.Description, vbCritical
    CleanupConnection
    OpenConnection = False
End Function

'@summary Checks if a connection to the database is currently open.
'@returns True if a connection is open, False otherwise.
Public Function IsConnected() As Boolean
    On Error Resume Next
    IsConnected = Not con Is Nothing And con.State = adStateOpen
    On Error GoTo 0
End Function

'@summary Closes the connection to the database.
Public Sub CloseConnection()
    If IsConnected Then
        If transactionInProgress Then RollbackTransaction
        con.Close
        Set con = Nothing
    End If
End Sub

' Returns the active connection
Public Function GetConnection() As ADODB.Connection
    If Not IsConnected Then
        MsgBox "Connection is not open.", vbCritical
        Set GetConnection = Nothing
    Else
        Set GetConnection = con
    End If
End Function

'@summary Executes a SQL query and returns the result as a recordset.
'@param sqlQuery The SQL query to execute.
'@returns The result of the query as a recordset.
Public Function ExecuteQuery(sqlQuery As String) As ADODB.Recordset
    Dim cmd As ADODB.Command, rs As ADODB.Recordset
    On Error GoTo QueryError

    Set cmd = CreateCommand(sqlQuery, GetConnection)
    Set rs = New ADODB.Recordset
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    Set ExecuteQuery = rs
    Exit Function

QueryError:
    MsgBox "Error executing query: " & Err.Description, vbCritical
    Set ExecuteQuery = Nothing
End Function

'@summary Creates a command object with the specified SQL text and connection.
'@param sqlText The SQL text for the command.
'@param connection The database connection to use.
'@returns The created command object.
Private Function CreateCommand(sqlText As String, connection As ADODB.Connection) As ADODB.Command
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = connection
    cmd.CommandText = sqlText
    cmd.CommandType = adCmdText
    Set CreateCommand = cmd
End Function


'' Transaction process - not yet in use
'@summary Begins a database transaction.
Public Sub BeginTransaction()
    If IsConnected And Not transactionInProgress Then
        con.BeginTrans
        transactionInProgress = True
    End If
End Sub

'@summary Commits the current
'@summary Commits the current database transaction.
Public Sub CommitTransaction()
    If IsConnected And transactionInProgress Then
        con.CommitTrans
        transactionInProgress = False
    End If
End Sub

'@summary Rolls back the current database transaction.
Public Sub RollbackTransaction()
    If IsConnected And transactionInProgress Then
        con.RollbackTrans
        transactionInProgress = False
    End If
End Sub