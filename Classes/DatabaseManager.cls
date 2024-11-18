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
Public Function ExecuteQuery(sqlText As String) As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    On Error GoTo QueryError

    ' Create and execute the command
    Set cmd = CreateCommand(sqlText, GetConnection)
    Set rs = New ADODB.Recordset
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    Set ExecuteQuery = rs
    Exit Function

QueryError:
    MsgBox "Error executing query: " & Err.Description & vbCrLf & "SQL: " & sqlText, vbCritical
    Set ExecuteQuery = Nothing
End Function

' For SQL calls with multiple tables
Sub ExecuteJoinQuery()
    Dim dbManager As New DatabaseManager
    Dim rs As ADODB.Recordset
    Dim query As String
    Dim columns As Variant
    Dim joins As Object
    Dim filters As Object

    ' Define the query components
    columns = Array("a.COLUMN_1", "b.COLUMN_2")
    Set joins = CreateObject("Scripting.Dictionary")
    joins.Add "TableB b", "a.ID = b.ID"
    Set filters = CreateObject("Scripting.Dictionary")
    filters.Add "a.Status", "Active"

    ' Build the query
    query = BuildJoinQuery("TableA a", joins, columns, filters)

    ' Execute the query
    If dbManager.OpenConnection Then
        Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(query, dbManager.GetConnection))
        If Not rs.EOF Then
            ' Skip metadata when populating data
            Call PopulateData(rs, False)
        End If
        rs.Close
    Else
        MsgBox "Failed to connect to the database.", vbCritical
    End If
End Sub


' Executes an action query (INSERT, UPDATE, DELETE).
Public Sub ExecuteNonQuery(sqlText As String)
    Dim cmd As ADODB.Command
    On Error GoTo NonQueryError

    ' Create and execute the command
    Set cmd = CreateCommand(sqlText, GetConnection)
    cmd.Execute
    Exit Sub

NonQueryError:
    MsgBox "Error executing non-query: " & Err.Description & vbCrLf & "SQL: " & sqlText, vbCritical
End Sub

' Creates a command object with the given SQL text.
Private Function CreateCommand(sqlText As String, connection As ADODB.Connection) As ADODB.Command
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = connection
    cmd.CommandText = sqlText
    cmd.CommandType = adCmdText
    Set CreateCommand = cmd
End Function

' Handles the cleanup of the connection object in case of unexpected errors.
Private Sub CleanupConnection()
    If con Is Nothing Then Exit Sub
    If con.State = adStateOpen Then con.Close
    Set con = Nothing
    transactionInProgress = False
End Sub
