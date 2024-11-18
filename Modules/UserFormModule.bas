' UserFormModule.bas
Option Explicit

' Opens the UserForm to collect user credentials.
Public Sub OpenUserForm()
    UserForm1.Show
End Sub

' UserFormModule
Sub LoginUser()
    Dim dbManager As New DatabaseManager

    ' Set credentials from user form
    dbManager.SetCredentials UserForm1.UserIDTextBox.Value, UserForm1.PasswordTextBox.Value

    ' Try to open connection
    If dbManager.OpenConnection Then
        MsgBox "Login successful!", vbInformation
        UserForm1.Hide
    Else
        MsgBox "Login failed! Please check your credentials.", vbCritical
    End If
End Sub

Sub LogoutUser()
    Dim dbManager As New DatabaseManager

    ' Close connection and notify user
    dbManager.CloseConnection
    MsgBox "You have been logged out.", vbInformation
End Sub
