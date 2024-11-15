' UserFormModule.bas
Option Explicit

Public Function GetUserCredentials() As Variant
    Dim userID As String
    Dim password As String
    
    ' Show the UserForm
    UserForm1.Show vbModal
    
    ' Retrieve the user ID and password
    userID = UserForm1.TextBoxUserID.Text
    password = UserForm1.TextBoxPassword.Text
    
    ' Return the credentials as an array
    GetUserCredentials = Array(userID, password)
    
    ' Unload the form
    Unload UserForm1
End Function

Public Sub ProcessUserCredentials()
    Dim credentials As Variant
    Dim userID As String
    Dim password As String
    Dim dbManager As DatabaseManager
    Dim userManager As UserManager
    
    ' Get user credentials
    credentials = GetUserCredentials()
    userID = credentials(0)
    password = credentials(1)
    
    ' Initialize UserManager
    Set userManager = New UserManager
    
    ' Validate user credentials
    If userManager.ValidateUser(userID) Then
        ' Initialize DatabaseManager
        Set dbManager = New DatabaseManager
        
        ' Open database connection
        If dbManager.OpenConnectionWithCredentials(password) Then
            MsgBox "Connection successful!", vbInformation
        Else
            MsgBox "Connection failed.", vbCritical
        End If
    Else
        MsgBox "Invalid user credentials.", vbCritical
    End If
End Sub