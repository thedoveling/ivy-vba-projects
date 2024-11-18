
' MainModule.bas
Option Explicit

Public Sub MainWorkflow()
    Dim dbManager As DatabaseManager
    Dim dataHandler As DataHandler
    Dim configManager As ConfigManager
    Dim rs As ADODB.Recordset
    Dim sqlQuery As String
    Dim userID As String
    Dim password As String
    Dim tableName As String
    
    ' Define your SQL query and table name
    tableName = "YOUR_TABLE_NAME"
    sqlQuery = "SELECT * FROM " & tableName & " WHERE ID = '124'"
    
    ' Initialize DatabaseManager
    Set dbManager = New DatabaseManager
    
    ' Open database connection
    If dbManager.OpenConnection Then
        ' Execute SQL query and get recordset
        Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(sqlQuery, dbManager.GetConnection))
        
        ' Initialize ConfigManager
        Set configManager = New ConfigManager
        configManager.Initialize tableName, "SELECT column_name AS variable, data_type AS datatype FROM all_tab_columns WHERE table_name = '" & tableName & "'", userID, password
        
        ' Initialize DataHandler
        Set dataHandler = New DataHandler
        
        ' Populate data as a table and apply configurations
        dataHandler.PopulateData rs
        
        ' Apply additional configurations (data validation and tooltips)
        Dim citedVariables As Collection
        Set citedVariables = New Collection
        Dim col As Long
        For col = 0 To rs.Fields.Count - 1
            citedVariables.Add rs.Fields(col).Name
        Next col
        dataHandler.ApplyConfigurations citedVariables
        
        ' Close recordset and connection
        rs.Close
        dbManager.CloseConnection
        
        MsgBox "Data populated and configured successfully!", vbInformation
    Else
        MsgBox "Connection failed.", vbCritical
    End If
End Sub

