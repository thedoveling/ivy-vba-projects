Option Explicit

' Opens the UserForm to collect user credentials.
Public Sub OpenUserForm()
    UserForm1.Show
End Sub

' Handles the login process when the user clicks the login button.
Sub LoginUser()
    Dim dbManager As DatabaseManager
    Set dbManager = New DatabaseManager

    ' Prompt user for credentials (e.g., from a UserForm)
    Dim userID As String, password As String
    ' Get user credentials from the UserForm
    userID = Trim(UserForm1.UserIDTextBox.Value)
    password = Trim(UserForm1.PasswordTextBox.Value)
    

    ' Set credentials and open the connection
    dbManager.SetCredentials userID, password
    If dbManager.OpenConnection Then
        MsgBox "Login successful!", vbInformation
        ' Close the UserForm
        UserForm1.Hide
    Else
        MsgBox "Login failed!", vbCritical
    End If
End Sub

Sub LogoutUser()
    Dim dbManager As DatabaseManager
    Set dbManager = New DatabaseManager

    dbManager.CloseConnection
    MsgBox "You have been logged out.", vbInformation
End Sub
