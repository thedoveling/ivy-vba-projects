' UserManager.cls
Option Explicit

Private dbManager As DatabaseManager

' Initializes the UserManager with a DatabaseManager instance
Public Sub Initialize(manager As DatabaseManager)
    Set dbManager = manager
End Sub

' Prompts user for credentials, validates, and opens a connection
Public Function Authenticate() As Boolean
    Dim userid As String
    Dim password As String
    Dim connectionString As String
    
    userid = InputBox("Enter Username:")
    password = InputBox("Enter Password:", "Password", , , , True)
    
    If userid = "" Or password = "" Then
        MsgBox "Username or password cannot be empty.", vbCritical
        Authenticate = False
        Exit Function
    End If
    
    ' Define the connection string based on provided credentials
    connectionString = "Provider=ORAOLEDB.ORACLE;Data Source=DEVVV;user id=" & Environ("UserName") & ";password=" & password & ";"
    
    ' Attempt to open a connection
    dbManager.OpenConnection connectionString
    Authenticate = dbManager.IsConnected
    If Not Authenticate Then
        MsgBox "Authentication failed. Please check credentials.", vbCritical
    End If
End Function

