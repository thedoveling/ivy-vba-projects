' UserManager.cls
Option Explicit

' Validates the user credentials.
' @param username - The username to validate
' @return - Boolean indicating success or failure
Public Function ValidateUser(username As String) As Boolean
    ' Example validation logic
    If username = Environ$("UserName") Then
        ValidateUser = True
    Else
        ValidateUser = False
    End If
End Function

' Retrieves the current system username.
' @return - The current system username
Public Function GetCurrentUsername() As String
    GetCurrentUsername = Environ$("UserName")
End Function

Sub LogoutUser()
    Dim dbManager As DatabaseManager
    Set dbManager = New DatabaseManager

    dbManager.CloseConnection
    MsgBox "You have been logged out.", vbInformation
End Sub
