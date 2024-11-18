' DatabaseManager.cls
Option Explicit

Private con As ADODB.Connection
Private userID As String
Private password As String

Private transactionInProgress As Boolean

' Sets user credentials.
' @param uID - The user ID
' @param pwd - The password
Public Sub SetCredentials(uID As String, pwd As String)
    userID = uID
    password = pwd
End Sub

' Opens a database connection using the globally stored credentials.
Public Function OpenConnection() As Boolean
    On Error GoTo ConnectionError

    ' Check if connection is already open
    If Not IsConnected Then
        ' Build and open the connection
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
    If Not con Is Nothing Then
        If con.State = adStateOpen Then con.Close
        Set con = Nothing
    End If
    OpenConnection = False
End Function

' Checks if the connection is active.
Public Function IsConnected() As Boolean
    On Error Resume Next ' Suppress errors during the check
    IsConnected = Not con Is Nothing And con.State = adStateOpen
    On Error GoTo 0 ' Restore normal error handling
End Function

' Closes the database connection.
Public Sub CloseConnection()
    If IsConnected Then
        If transactionInProgress Then
            MsgBox "Warning: Closing connection with an active transaction. Rolling back changes.", vbExclamation
            RollbackTransaction
        End If
        con.Close
        Set con = Nothing
    End If
End Sub

' Begins a transaction.
Public Sub BeginTransaction()
    If IsConnected Then
        If transactionInProgress Then
            MsgBox "Transaction already in progress.", vbExclamation
        Else
            con.BeginTrans
            transactionInProgress = True
        End If
    Else
        MsgBox "Cannot start a transaction. Database connection is not open.", vbCritical
    End If
End Sub

' Commits the current transaction.
Public Sub CommitTransaction()
    If IsConnected And transactionInProgress Then
        con.CommitTrans
        transactionInProgress = False
    ElseIf Not transactionInProgress Then
        MsgBox "No transaction to commit.", vbExclamation
    Else
        MsgBox "Cannot commit. Database connection is not open.", vbCritical
    End If
End Sub

' Rolls back the current transaction.
Public Sub RollbackTransaction()
    If IsConnected And transactionInProgress Then
        con.RollbackTrans
        transactionInProgress = False
    ElseIf Not transactionInProgress Then
        MsgBox "No transaction to roll back.", vbExclamation
    Else
        MsgBox "Cannot roll back. Database connection is not open.", vbCritical
    End If
End Sub

' Returns the active connection.
Public Function GetConnection() As ADODB.Connection
    If Not IsConnected Then
        MsgBox "Database connection is not open.", vbCritical
        Set GetConnection = Nothing
    Else
        Set GetConnection = con
    End If
End Function

' Executes a SELECT query and returns a recordset.
Public Function ExecuteQuery(sqlQuery As String) As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    On Error GoTo QueryError

    ' Create and execute the command
    Set cmd = CreateCommand(sqlQuery, GetConnection)
    Set rs = New ADODB.Recordset
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    Set ExecuteQuery = rs
    Exit Function

QueryError:
    MsgBox "Error executing query: " & Err.Description & vbCrLf & "SQL: " & sqlText, vbCritical
    Set ExecuteQuery = Nothing
End Function


' Executes a JOIN query and returns a recordset.
Public Function ExecuteJoinQuery(joinquery As String) As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    On Error GoTo JoinQueryError

    ' Create and execute the command
    Set cmd = CreateCommand(joinquery, GetConnection)
    Set rs = New ADODB.Recordset
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    Set ExecuteJoinQuery = rs
    Exit Function

JoinQueryError:
    MsgBox "Error executing JOIN query: " & Err.Description & vbCrLf & "SQL: " & sqlText, vbCritical
    Set ExecuteJoinQuery = Nothing
End Function

' Handles the cleanup of the connection object in case of unexpected errors.
Private Sub CleanupConnection()
    If con Is Nothing Then Exit Sub
    If con.State = adStateOpen Then con.Close
    Set con = Nothing
    transactionInProgress = False
End Sub
