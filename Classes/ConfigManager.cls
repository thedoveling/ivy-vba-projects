Sub ExampleWorkflow()
    Dim targetSheet As Worksheet
    Dim rs As ADODB.Recordset
    Dim dbManager As DatabaseManager
    Dim logFile As String
    Dim dataHandler As DataHandler
    
    ' Initialize objects
    Set targetSheet = ThisWorkbook.Sheets("Data")
    Set dbManager = New DatabaseManager
    Set dataHandler = New DataHandler
    logFile = ThisWorkbook.Path & "\error_log.txt"
    
    On Error GoTo ErrorHandler
    
    ' Disable events and screen updating
    Call ErrorHandler.DisableEvents
    
    ' Open database connection
    If dbManager.OpenConnectionWithCredentials("password") Then
        ' Build and execute query
        Set rs = dbManager.ExecuteQuery("SELECT * FROM YourTable WHERE Condition = 'Value'")
        
        ' Populate data with dynamic headers
        Call dataHandler.PopulateData(targetSheet, rs)
        
        ' Apply configurations from the codebook
        Call dataHandler.ApplyConfigurations(targetSheet)
        
        ' Close recordset and connection
        rs.Close
        dbManager.CloseConnection
    Else
        Call ErrorHandler.HandleError("Failed to open database connection.", logFile)
    End If
    
    ' Enable events and screen updating
    Call ErrorHandler.EnableEvents
    
    Exit Sub
    
ErrorHandler:
    ' Handle runtime error
    Call ErrorHandler.HandleRuntimeError(logFile)
    ' Enable events and screen updating
    Call ErrorHandler.EnableEvents
    Resume Next
End Sub