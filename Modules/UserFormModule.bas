Option Explicit

' Shows the user form and retrieves credentials.
Public Function GetUserCredentials() As Variant
    Dim userID As String, password As String

    On Error GoTo ErrorHandler
    UserForm1.Show vbModal
    
    userID = UserForm1.TextBoxUserID.Text
    password = UserForm1.TextBoxPassword.Text

    GetUserCredentials = Array(userID, password)
    Unload UserForm1
    Exit Function

ErrorHandler:
    Call HandleRuntimeError("Error in GetUserCredentials")
    GetUserCredentials = Array("", "") ' Return empty credentials on error
End Function

' Processes user credentials for authentication and database connection.
Public Sub ProcessUserCredentials()
    Dim credentials As Variant
    Dim userID As String, password As String
    Dim dbManager As DatabaseManager, userManager As UserManager

    On Error GoTo ErrorHandler

    credentials = GetUserCredentials()
    userID = credentials(0)
    password = credentials(1)

    If Trim(userID) = "" Or Trim(password) = "" Then
        MsgBox "User ID or password cannot be empty.", vbExclamation
        Exit Sub
    End If

    ' Validate User
    Set userManager = New UserManager
    If Not userManager.ValidateUser(userID) Then
        MsgBox "Invalid User ID.", vbCritical
        Exit Sub
    End If

    ' Connect to Database
    Set dbManager = New DatabaseManager
    If dbManager.OpenConnectionWithCredentials(password) Then
        MsgBox "Connection successful!", vbInformation
    Else
        MsgBox "Connection failed.", vbCritical
    End If

    Exit Sub

ErrorHandler:
    Call HandleRuntimeError("Error in ProcessUserCredentials")
End Sub
