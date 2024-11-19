Sub SQLQueryFetch()
    Dim dbManager As DatabaseManager
    Dim configManager As ConfigManager
    Dim dataHandler As DataHandler
    Dim rs As ADODB.Recordset
    Dim tableName As String

    ' Step 1: Initialize DatabaseManager
    Set dbManager = New DatabaseManager

    If Not dbManager.OpenConnection Then
        MsgBox "Database connection failed.", vbCritical
        Exit Sub
    End If

    ' Step 2: Set table name and initialize ConfigManager
    tableName = "YOUR_TABLE_NAME"
    Set configManager = New ConfigManager
    configManager.Initialize tableName, dbManager ' Corrected order of parameters

    ' Step 3: Fetch data using DatabaseManager
    Set rs = dbManager.ExecuteQuery(SQLHelper.BuildSelectQuery(tableName))

    ' Step 4: Populate data using DataHandler
    Set dataHandler = New DataHandler
    dataHandler.PopulateData rs, True, configManager

    ' Step 5: Clean up
    rs.Close
    dbManager.CloseConnection
End Sub
