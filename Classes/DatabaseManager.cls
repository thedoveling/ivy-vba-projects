' DatabaseManager.cls
Option Explicit

Private con As ADODB.connection

' Opens a database connection using provided credentials.
' @param the userid in the system environment - The user ID for the connection
' @param password - The password for the connection
' @return - Boolean indicating success or failure
' Opens a connection with userid and password credentials

Public Function OpenConnectionWithCredentials(password As String) As Boolean
    On Error GoTo ConnectionError
    Dim connectionString As String
    
    ' Define the connection string with userid and password
    connectionString = "Provider=ORAOLEDB.ORACLE;Data Source=DEVVV;user id=" & Environ$("UserName") & ";" & "password=" & password & ";"
    Set con = New ADODB.connection
    
    con.connectionString = connectionString
    con.Open
    OpenConnectionWithCredentials = True
    Exit Function

ConnectionError:
    MsgBox "Error opening connection: " & Err.Description, vbCritical
    OpenConnectionWithCredentials = False
End Function

' Closes the database connection if open.
Public Sub CloseConnection()
    If Not con Is Nothing Then
        If con.State = adStateOpen Then
            con.Close
        End If
        Set con = Nothing
    End If
End Sub

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

' Executes a parameterized query within a transaction.
' query: SQL query string.
' params: Array of ADODB.Parameters for the query.
' Returns: ADODB.Recordset or Nothing if not a select query.
Public Function ExecuteQueryWithTransaction(query As String, ParamArray params() As Variant) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandText = query
    
    ' Append parameters
    Dim i As Integer
    For i = LBound(params) To UBound(params)
        cmd.Parameters.Append params(i)
    Next i
    
    ' Execute the query and return result
    Set ExecuteQueryWithTransaction = cmd.Execute
    Exit Function

ErrorHandler:
    ' Rollback transaction and handle the error
    Call RollbackTransaction
    ErrorEventManager.HandleError Err.Description, "ExecuteQueryWithTransaction"
    Set ExecuteQueryWithTransaction = Nothing
End Function

