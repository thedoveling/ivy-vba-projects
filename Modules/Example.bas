Sub ExampleUsage()
    Dim credentials As Variant
    Dim userID As String
    Dim password As String
    Dim dbManager As DatabaseManager
    Dim userManager As UserManager
    Dim dataHandler As DataHandler
    Dim rs As ADODB.Recordset
    Dim targetSheet As Worksheet
    
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
            ' Execute SQL query and get recordset
            Set rs = dbManager.ExecuteQuery("SELECT field1, field2, field3 FROM YourTable WHERE Condition = 'Value'")
            
            ' Initialize DataHandler
            Set dataHandler = New DataHandler
            
            ' Set target sheet
            Set targetSheet = ThisWorkbook.Sheets("Data")
            
            ' Populate data as a table
            dataHandler.PopulateData targetSheet, rs
            
            ' Close recordset and connection
            rs.Close
            dbManager.CloseConnection
            
            MsgBox "Data populated successfully!", vbInformation
        Else
            MsgBox "Connection failed.", vbCritical
        End If
    Else
        MsgBox "Invalid user credentials.", vbCritical
    End If
End Sub