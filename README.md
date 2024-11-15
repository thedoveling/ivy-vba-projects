# ivy-vba-projects
 Development for VBA projects. 


Example of use: 
Sub ExampleWorkflow()
    Dim targetSheet As Worksheet
    Dim rs As ADODB.Recordset
    Dim dbManager As DatabaseManager
    Dim logFile As String
    Dim dataHandler As DataHandler
    Dim citedVariables As Collection
    Dim sqlQuery As String
    Dim variable As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' Initialize objects
    Set targetSheet = ThisWorkbook.Sheets("Data")
    Set dbManager = New DatabaseManager
    Set dataHandler = New DataHandler
    Set citedVariables = New Collection
    logFile = ThisWorkbook.Path & "\error_log.txt"
    
    On Error GoTo ErrorHandler
    
    ' Disable events and screen updating
    Call ErrorHandler.DisableEvents
    
    ' Define SQL query
    sqlQuery = "SELECT field1, field2, field3 FROM YourTable WHERE Condition = 'Value'"
    
    ' Extract cited variables from SQL query
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "\b\w+\b"
    Set matches = regex.Execute(sqlQuery)
    
    For Each match In matches
        variable = match.Value
        If Not IsInCollection(citedVariables, variable) Then
            citedVariables.Add variable
        End If
    Next match
    
    ' Open database connection
    If dbManager.OpenConnectionWithCredentials("password") Then
        ' Execute query
        Set rs = dbManager.ExecuteQuery(sqlQuery)
        
        ' Populate data with dynamic headers
        Call dataHandler.PopulateData(targetSheet, rs)
        
        ' Apply configurations from the codebook
        Call dataHandler.ApplyConfigurations(targetSheet, citedVariables)
        
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

' Checks if an item is in a collection
' @param col - The collection
' @param item - The item to check
' @return - Boolean indicating if the item is in the collection
Function IsInCollection(col As Collection, item As Variant) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = col(item)
    IsInCollection = (Err.Number = 0)
    On Error GoTo 0
End Function