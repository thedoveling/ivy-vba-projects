' UnitTest_ConfigManager.bas
Option Explicit

Public Sub Test_ConfigManager()
    Dim configManager As ConfigManager
    Dim columnMappings As Object, dataValidationConfigs As Object
    Dim fieldOptions As Object, tooltips As Object
    Dim tableName As String
    Dim sqlQuery As String

    ' Define the table name and SQL query
    tableName = "YOUR_TABLE_NAME"
    sqlQuery = "SELECT column_name AS variable, data_type AS datatype FROM all_tab_columns WHERE table_name = '" & tableName & "'"
    
    ' Initialize ConfigManager
    Set configManager = New ConfigManager
    configManager.Initialize tableName, sqlQuery
    ' Fetch dynamically loaded data
    Set columnMappings = configManager.GetColumnMappings()
    Set dataValidationConfigs = configManager.GetDataValidationConfigs()
    Set fieldOptions = configManager.GetFieldOptions()
    Set tooltips = configManager.GetTooltips()

    ' Assertions (mock data checks)
    Debug.Assert Not columnMappings Is Nothing
    Debug.Assert Not dataValidationConfigs Is Nothing
    Debug.Assert Not fieldOptions Is Nothing
    Debug.Assert Not tooltips Is Nothing

    ' Check for specific mock values
    Debug.Assert columnMappings.Exists("variable") ' Check Oracle-loaded variable
    Debug.Assert fieldOptions.Exists("variable")    ' Check local Config variable
    Debug.Print "Test_ConfigManager passed."
End Sub


' TestModule.bas
Option Explicit

Public Sub TestOracleConnection()
    Dim dbManager As DatabaseManager
    Dim rs As ADODB.Recordset
    Dim sqlQuery As String
    
    ' Define your SQL query to pull one record
    sqlQuery = "SELECT * FROM YOUR_TABLE_NAME WHERE ID = '124'"
    
    ' Initialize DatabaseManager
    Set dbManager = New DatabaseManager
    
    ' Open database connection
    If dbManager.OpenConnectionWithCredentials("password") Then
        ' Execute SQL query and get recordset
        Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(sqlQuery, dbManager.GetConnection))
        
        ' Print the result to the Immediate Window
        If Not rs.EOF Then
            Debug.Print "ID: " & rs.Fields("ID").Value
            Debug.Print "OtherField: " & rs.Fields("OtherField").Value ' Replace with actual field names
        Else
            Debug.Print "No records found."
        End If
        
        ' Close recordset and connection
        rs.Close
        dbManager.CloseConnection
    Else
        Debug.Print "Connection failed."
    End If
End Sub


' TestModule.bas
Option Explicit

Public Sub TestOracleConnection()
    Dim dbManager As DatabaseManager
    Dim rs As ADODB.Recordset
    Dim sqlQuery As String
    Dim userID As String
    Dim password As String
    
    ' Define your SQL query to pull one record
    sqlQuery = "SELECT * FROM YOUR_TABLE_NAME WHERE ID = '124'"
    
    ' Get user credentials (assuming you have a method to get these)
    userID = "your_user_id" ' Replace with actual user ID
    password = "your_password" ' Replace with actual password
    
    ' Initialize DatabaseManager
    Set dbManager = New DatabaseManager
    
    ' Debug statements to verify credentials
    Debug.Print "User ID: " & userID
    Debug.Print "Password: " & password
    
    ' Open database connection
    If dbManager.OpenConnectionWithCredentials(userID, password) Then
        ' Execute SQL query and get recordset
        Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(sqlQuery, dbManager.GetConnection))
        
        ' Print the result to the Immediate Window
        If Not rs.EOF Then
            Debug.Print "ID: " & rs.Fields("ID").Value
            Debug.Print "OtherField: " & rs.Fields("OtherField").Value ' Replace with actual field names
        Else
            Debug.Print "No records found."
        End If
        
        ' Close recordset and connection
        rs.Close
        dbManager.CloseConnection
    Else
        Debug.Print "Connection failed."
    End If
End Sub