' ConfigManager.bas
Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary

' Initializes the ConfigManager by loading configurations.
Public Sub Initialize()
    Set columnMappings = New Scripting.Dictionary
    Set dataValidationConfigs = New Scripting.Dictionary
    
    ' Load configurations from a configuration sheet (optional)
    ' Call LoadConfigurationsFromSheet(ThisWorkbook.Sheets("ConfigSheet"))
End Sub

' Maps column headers to Oracle fields.
' @return - A dictionary of column mappings
Public Function GetColumnMappings() As Scripting.Dictionary
    Set GetColumnMappings = columnMappings
End Function

' Adds a column mapping.
' @param header - The column header
' @param oracleField - The corresponding Oracle field
Public Sub AddColumnMapping(header As String, oracleField As String)
    If Not columnMappings.Exists(header) Then
        columnMappings.Add header, oracleField
    End If
End Sub

' Retrieves data validation configurations.
' @return - A dictionary of data validation configurations
Public Function GetDataValidationConfigs() As Scripting.Dictionary
    Set GetDataValidationConfigs = dataValidationConfigs
End Function

' Adds a data validation configuration.
' @param column - The column name
' @param validationRule - The validation rule
Public Sub AddDataValidationConfig(column As String, validationRule As String)
    If Not dataValidationConfigs.Exists(column) Then
        dataValidationConfigs.Add column, validationRule
    End If
End Sub

' Loads configurations from a configuration sheet.
' @param configSheet - The configuration sheet
Public Sub LoadConfigurationsFromSheet(configSheet As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim header As String
    Dim oracleField As String
    Dim column As String
    Dim validationRule As String
    
    lastRow = configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).Row
    
    ' Load column mappings
    For i = 2 To lastRow
        header = configSheet.Cells(i, 1).Value
        oracleField = configSheet.Cells(i, 2).Value
        If header <> "" And oracleField <> "" Then
            Call AddColumnMapping(header, oracleField)
        End If
    Next i
    
    ' Load data validation configurations
    For i = 2 To lastRow
        column = configSheet.Cells(i, 3).Value
        validationRule = configSheet.Cells(i, 4).Value
        If column <> "" And validationRule <> "" Then
            Call AddDataValidationConfig(column, validationRule)
        End If
    Next i
End Sub