Option Explicit

' Opens the UserForm to collect user credentials.
Public Sub OpenUserForm()
    UserForm1.Show
End Sub

' Handles the login process when the user clicks the login button.
Sub LoginUser()
    Dim dbManager As DatabaseManager
    Set dbManager = New DatabaseManager
    
    ' Open the connection and check if login is successful
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
