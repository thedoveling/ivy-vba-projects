' DatabaseManager.cls
Option Explicit

Private con As ADODB.Connection

' Opens a database connection using the system environment user ID and provided password.
' @param password - The password for the connection
' @return - Boolean indicating success or failure
Public Function OpenConnectionWithCredentials(password As String) As Boolean
    On Error GoTo ConnectionError
    Dim connectionString As String
    
    ' Define the connection string with user ID from environment and provided password
    connectionString = "Provider=ORAOLEDB.ORACLE;Data Source=DEVVV;User Id=" & Environ$("UserName") & ";Password=" & password & ";"
    Set con = New ADODB.Connection
    
    con.ConnectionString = connectionString
    con.Open
    OpenConnectionWithCredentials = True
    Exit Function

ConnectionError:
    MsgBox "Error opening connection: " & Err.Description, vbCritical
    If Not con Is Nothing Then
        If con.State = adStateOpen Then
            con.Close
        End If
        Set con = Nothing
    End If
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

' Executes a SQL query and returns the result as a recordset.
' @param sqlQuery - The SQL query to execute
' @return - ADODB.Recordset containing the query result
Public Function ExecuteQuery(sqlQuery As String) As ADODB.Recordset
    On Error GoTo QueryError
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open sqlQuery, con, adOpenStatic, adLockReadOnly
    Set ExecuteQuery = rs
    Exit Function

QueryError:
    MsgBox "Error executing query: " & Err.Description, vbCritical
    Set ExecuteQuery = Nothing
End Function

' Creates a parameter for a command.
' @param cmd - The command object
' @param paramName - The name of the parameter
' @param paramType - The type of the parameter
' @param paramValue - The value of the parameter
Public Sub CreateParameter(cmd As ADODB.Command, paramName As String, paramType As ADODB.DataTypeEnum, paramValue As Variant)
    Dim param As ADODB.Parameter
    Set param = cmd.CreateParameter(paramName, paramType, adParamInput, , paramValue)
    cmd.Parameters.Append param
End Sub