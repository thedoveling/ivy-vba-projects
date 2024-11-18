' ConfigManager.cls
Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary
Private fieldOptions As Scripting.Dictionary
Private tooltips As Scripting.Dictionary

Private dbManager As DatabaseManager

' Initializes the ConfigManager by loading configurations from Oracle and local configuration file.
' @param tableName - The name of the table to load configurations for
' @param sqlQuery - The SQL query to execute for loading configurations
Public Sub Initialize(tableName As String, sqlQuery As String, dbManager As DatabaseManager)
    Dim metadataQuery As String
    Dim rs As ADODB.Recordset

    Set columnMappings = New Scripting.Dictionary
    Set dataValidationConfigs = New Scripting.Dictionary
    Set fieldOptions = New Scripting.Dictionary
    Set tooltips = New Scripting.Dictionary
    
    ' Set the DatabaseManager instance
    Set dbManager = New DatabaseManager

    ' Build a default SQL query if one is not provided
    If sqlQuery = "" Then
        metadataQuery = "SELECT COLUMN_NAME, DATA_TYPE FROM USER_TAB_COLUMNS WHERE TABLE_NAME = '" & tableName & "'"
    Else
        metadataQuery = sqlQuery
    End If
    
    ' Open database connection
    If dbManager.OpenConnection Then
        ' Load configurations from Oracle
        Call LoadConfigurationsFromOracle(tableName)
        
        ' Load field options and tooltips from local configuration file
        Call LoadLocalConfigurations
        
        ' Close database connection
        dbManager.CloseConnection
    Else
        MsgBox "Failed to open database connection.", vbCritical
    End If
End Sub

' Maps column headers to Oracle fields.
' @return - A dictionary of column mappings
Public Function GetColumnMappings() As Scripting.Dictionary
    Set GetColumnMappings = columnMappings
End Function

' Retrieves data validation configurations.
' @return - A dictionary of data validation configurations
Public Function GetDataValidationConfigs() As Scripting.Dictionary
    Set GetDataValidationConfigs = dataValidationConfigs
End Function

' Retrieves field options for drop-downs.
' @return - A dictionary of field options
Public Function GetFieldOptions() As Scripting.Dictionary
    Set GetFieldOptions = fieldOptions
End Function

' Retrieves tooltips for headers.
' @return - A dictionary of tooltips
Public Function GetTooltips() As Scripting.Dictionary
    Set GetTooltips = tooltips
End Function

' Loads configurations from Oracle.
' @param tableName - The name of the table to load configurations for
' @param sqlQuery - The SQL query to execute for loading configurations

' Dynamically builds and executes the metadata query to load configurations from Oracle.
' @param tableName - The name of the table to load configurations for
Private Sub LoadConfigurationsFromOracle(tableName As String)
    Dim rs As ADODB.Recordset
    Dim columnName As String
    Dim dataType As String
    Dim metadataQuery As String

    On Error GoTo LoadError

    ' Dynamically build the query
    metadataQuery = "SELECT COLUMN_NAME, DATA_TYPE FROM USER_TAB_COLUMNS WHERE TABLE_NAME = '" & tableName & "'"

    ' Execute the query
    Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(metadataQuery, dbManager.GetConnection))

    ' Process the recordset to populate mappings and configurations
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            columnName = rs.Fields("COLUMN_NAME").Value
            dataType = rs.Fields("DATA_TYPE").Value

            ' Populate column mappings
            If Not columnMappings.Exists(columnName) Then columnMappings.Add columnName, columnName ' Default: column name is the friendly name

            ' Populate data validation configurations
            If Not dataValidationConfigs.Exists(columnName) Then dataValidationConfigs.Add columnName, dataType ' Use data type for validation rules

            rs.MoveNext
        Loop
        rs.Close
    End If

    Exit Sub

LoadError:
    Call HandleRuntimeError("Error loading configurations from Oracle")
End Sub

' Load local configurations from the Config worksheet.
Private Sub LoadLocalConfigurations()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim variable As String, options As String, tooltip As String
    Dim headerRow As Range
    Dim colVariable As Long, colOptions As Long, colTooltip As Long

    ' Set the Config worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find the header row
    Set headerRow = ws.ListObjects("Config").HeaderRowRange

    ' Map the columns dynamically
    colVariable = Application.Match("variable", headerRow, 0)
    colOptions = Application.Match("options", headerRow, 0)
    colTooltip = Application.Match("tooltip", headerRow, 0)

    ' Loop through the Config sheet to populate configurations
    For i = 2 To lastRow
        variable = ws.Cells(i, colVariable).Value
        options = ws.Cells(i, colOptions).Value
        tooltip = ws.Cells(i, colTooltip).Value

        ' Populate dictionaries dynamically
        If Not fieldOptions.Exists(variable) Then fieldOptions(variable) = options
        If Not tooltips.Exists(variable) Then tooltips(variable) = tooltip
    Next i
End Sub