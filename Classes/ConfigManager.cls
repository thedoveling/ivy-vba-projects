' ConfigManager.cls
Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary
Private fieldOptions As Scripting.Dictionary
Private tooltips As Scripting.Dictionary

' Initializes the ConfigManager by loading configurations.
Public Sub Initialize()
    Set columnMappings = New Scripting.Dictionary
    Set dataValidationConfigs = New Scripting.Dictionary
    Set fieldOptions = New Scripting.Dictionary
    Set tooltips = New Scripting.Dictionary
    
    ' Load configurations from the codebook sheet
    Call LoadConfigurationsFromCodebook(ThisWorkbook.Sheets("Codebook"))
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

' Loads configurations from the codebook sheet.
' @param codebookSheet - The codebook sheet
Public Sub LoadConfigurationsFromCodebook(codebookSheet As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim headers As Scripting.Dictionary
    Dim header As String
    Dim variable As String
    Dim datatype As String
    Dim validation As String
    Dim fieldOptionsStr As String
    Dim nullable As String
    Dim schema_name As String
    Dim server_name As String
    Dim system_name As String
    Dim tooltip As String
    
    On Error GoTo LoadError
    
    Set headers = New Scripting.Dictionary
    
    lastRow = codebookSheet.Cells(codebookSheet.Rows.Count, 1).End(xlUp).Row
    lastCol = codebookSheet.Cells(1, codebookSheet.Columns.Count).End(xlToLeft).Column
    
    ' Read headers from the first row
    For j = 1 To lastCol
        headers.Add codebookSheet.Cells(1, j).Value, j
    Next j
    
    ' Load configurations from the codebook
    For i = 2 To lastRow
        header = codebookSheet.Cells(i, headers("field_label")).Value
        variable = codebookSheet.Cells(i, headers("variable")).Value
        datatype = codebookSheet.Cells(i, headers("datatype")).Value
        validation = codebookSheet.Cells(i, headers("validation")).Value
        fieldOptionsStr = codebookSheet.Cells(i, headers("field_options")).Value
        nullable = codebookSheet.Cells(i, headers("nullable")).Value
        schema_name = codebookSheet.Cells(i, headers("schema_name")).Value
        server_name = codebookSheet.Cells(i, headers("server_name")).Value
        system_name = codebookSheet.Cells(i, headers("system_name")).Value
        tooltip = codebookSheet.Cells(i, headers("tooltip")).Value
        
        ' Add to column mappings
        If header <> "" And variable <> "" Then
            columnMappings.Add variable, header
        End If
        
        ' Add to data validation configurations
        If variable <> "" And validation <> "" Then
            dataValidationConfigs.Add variable, validation
        End If
        
        ' Add to field options
        If variable <> "" And fieldOptionsStr <> "" Then
            fieldOptions.Add variable, fieldOptionsStr
        End If
        
        ' Add to tooltips
        If variable <> "" And tooltip <> "" Then
            tooltips.Add variable, tooltip
        End If
    Next i
    
    Exit Sub

LoadError:
    MsgBox "Error loading configurations from codebook: " & Err.Description, vbCritical
End Sub