' DataHandler.cls
Option Explicit

' Populates data from a recordset into the specified Excel sheet.
' Dynamically sets headers based on the recordset fields.
' @param targetSheet - The Excel sheet to populate data in
' @param rs - The recordset containing the data
' @param headerRow - The row number to set headers in (default is 4)
' @param startRow - The row number to start populating data (default is 5)
Public Sub PopulateData(targetSheet As Worksheet, rs As ADODB.Recordset, Optional headerRow As Long = 4, Optional startRow As Long = 5)
    Dim col As Long
    
    ' Set headers dynamically based on recordset fields
    For col = 0 To rs.Fields.Count - 1
        targetSheet.Cells(headerRow, col + 1).Value = rs.Fields(col).Name
    Next col
    
    ' Populate data starting from startRow
    targetSheet.Cells(startRow, 1).CopyFromRecordset rs
End Sub

' Clears data in the specified range without affecting headers.
' @param targetSheet - The Excel sheet to clear data in
' @param dataRange - The range to clear data in
' @param headerRow - The row number where headers are located (default is 4)
Public Sub ClearData(targetSheet As Worksheet, dataRange As Range, Optional headerRow As Long = 4)
    Dim firstDataRow As Long
    firstDataRow = headerRow + 1
    dataRange.Rows(firstDataRow & ":" & dataRange.Rows.Count).ClearContents
End Sub

' Applies configurations from the codebook to the target sheet based on the variables cited in the SQL string.
' @param targetSheet - The Excel sheet to apply configurations to
' @param citedVariables - A collection of variables cited in the SQL string
Public Sub ApplyConfigurations(targetSheet As Worksheet, citedVariables As Collection)
    Dim configManager As ConfigManager
    Dim variable As String
    
    ' Initialize ConfigManager and load configurations
    Set configManager = New ConfigManager
    configManager.Initialize
    
    ' Apply configurations for each cited variable
    For Each variable In citedVariables
        ApplyTooltip targetSheet, configManager, variable
        ApplyFieldOptions targetSheet, configManager, variable
        ApplyDataValidation targetSheet, configManager, variable
    Next variable
End Sub

' Applies tooltip for a variable
' @param targetSheet - The Excel sheet to apply the tooltip to
' @param configManager - The ConfigManager instance
' @param variable - The variable to apply the tooltip for
Private Sub ApplyTooltip(targetSheet As Worksheet, configManager As ConfigManager, variable As String)
    Dim tooltips As Scripting.Dictionary
    Dim header As String
    Dim col As Long
    
    Set tooltips = configManager.GetTooltips
    If tooltips.Exists(variable) Then
        header = configManager.GetColumnMappings(variable)
        col = Application.Match(header, targetSheet.Rows(4), 0)
        If Not IsError(col) Then
            targetSheet.Cells(4, col).AddComment tooltips(variable)
        End If
    End If
End Sub

' Applies field options for a variable
' @param targetSheet - The Excel sheet to apply the field options to
' @param configManager - The ConfigManager instance
' @param variable - The variable to apply the field options for
Private Sub ApplyFieldOptions(targetSheet As Worksheet, configManager As ConfigManager, variable As String)
    Dim fieldOptions As Scripting.Dictionary
    Dim header As String
    Dim col As Long
    
    Set fieldOptions = configManager.GetFieldOptions
    If fieldOptions.Exists(variable) Then
        header = configManager.GetColumnMappings(variable)
        col = Application.Match(header, targetSheet.Rows(4), 0)
        If Not IsError(col) Then
            With targetSheet.Cells(5, col).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:=fieldOptions(variable)
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        End If
    End If
End Sub

' Applies data validation for a variable
' @param targetSheet - The Excel sheet to apply the data validation to
' @param configManager - The ConfigManager instance
' @param variable - The variable to apply the data validation for
Private Sub ApplyDataValidation(targetSheet As Worksheet, configManager As ConfigManager, variable As String)
    Dim dataValidationConfigs As Scripting.Dictionary
    Dim header As String
    Dim col As Long
    
    Set dataValidationConfigs = configManager.GetDataValidationConfigs
    If dataValidationConfigs.Exists(variable) Then
        header = configManager.GetColumnMappings(variable)
        col = Application.Match(header, targetSheet.Rows(4), 0)
        If Not IsError(col) Then
            ' Apply validation rule (example: numeric validation)
            If dataValidationConfigs(variable) = "numeric" Then
                With targetSheet.Cells(5, col).Validation
                    .Delete
                    .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="1", Formula2:="100"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                End With
            End If
        End If
    End If
End Sub