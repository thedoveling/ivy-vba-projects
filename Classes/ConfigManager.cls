' ConfigManager.cls
Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary
Private fieldOptions As Scripting.Dictionary
Private tooltips As Scripting.Dictionary

Private dbManager As DatabaseManager

' Initializes the configuration manager with an open database connection.
' @param tableName - The name of the table to load configurations for
' @param dbManager - The database manager object
Public Sub Initialize(tableName As String, dbManager As DatabaseManager)
    If Not dbManager.IsConnected Then
        MsgBox "Database connection is not open. Please establish a connection first.", vbCritical
        Exit Sub
    End If

    Set ConfigManager.dbManager = dbManager
    LoadConfigurationsFromOracle(tableName)
    LoadLocalConfigurations()
End Sub

' Loads configurations from Oracle using the database manager.
Private Sub LoadConfigurationsFromOracle(tableName As String)
    Dim rs As ADODB.Recordset
    Dim metadataQuery As String

    metadataQuery = "SELECT COLUMN_NAME, DATA_TYPE FROM ALL_TAB_COLUMNS WHERE OWNER IN ('ZZZHSO','ZZZHSF') AND TABLE_NAME = '" & tableName & "'"
    Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(metadataQuery, dbManager.GetConnection))

    Do While Not rs.EOF
        columnMappings.Add rs.Fields("COLUMN_NAME").Value, rs.Fields("COLUMN_NAME").Value
        dataValidationConfigs.Add rs.Fields("COLUMN_NAME").Value, rs.Fields("DATA_TYPE").Value
        rs.MoveNext
    Loop

    rs.Close
End Sub

' Retrieves column mappings.
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

' Loads local configurations from the Config worksheet.
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
