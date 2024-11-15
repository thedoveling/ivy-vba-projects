' Module1.bas
Option Explicit

Sub ExampleUsage()
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
    If userManager.ValidateUser(userID, password) Then
        ' Initialize DatabaseManager
        Set dbManager = New DatabaseManager
        
        ' Open database connection
        If dbManager.OpenConnectionWithCredentials(userID, password) Then
            MsgBox "Connection successful!", vbInformation
        Else
            MsgBox "Connection failed.", vbCritical
        End If
    Else
        MsgBox "Invalid user credentials.", vbCritical
    End If
End Sub