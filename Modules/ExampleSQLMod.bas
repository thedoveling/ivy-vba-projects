Sub SQLQueryFetch()
    Dim dbManager As DatabaseManager
    Dim configManager As ConfigManager
    Dim dataHandler As DataHandler
    Dim rs As ADODB.Recordset
    Dim tableName As String
    Dim schema As String

    ' Step 1: Initialize DatabaseManager
    Set dbManager = New DatabaseManager

    If Not dbManager.OpenConnection Then
        MsgBox "Database connection failed.", vbCritical
        Exit Sub
    End If

    ' Step 2: Set table name and initialize ConfigManager
    tableName = "YOUR_TABLE_NAME"
    schema = "zzzivy"
    Set configManager = New ConfigManager
    configManager.Initialize tableName, schema, dbManager ' Corrected order of parameters

    ' Step 3: Fetch metadata using ConfigManager
    Set rsMetadata = dbManager.ExecuteQuery(SQLHelper.BuildMetadataQuery(tableName, schema))

    ' Step 4: Fetch actual data using DatabaseManager
    Set rsData = dbManager.ExecuteQuery(SQLHelper.BuildSelectQuery(tableName, schema))

    ' Step 5: Populate data using DataHandler
    Set dataHandler = New DataHandler
    ' Pass metadata recordset for headers, and data recordset for rows
    dataHandler.PopulateData rsData, rsMetadata, configManager

    ' Step 6: Clean up
    rsMetadata.Close
    rsData.Close
    ' Step 5: Clean up
    rs.Close
    dbManager.CloseConnection
End Sub
