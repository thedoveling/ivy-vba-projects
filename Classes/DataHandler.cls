Option Explicit

Private Const DATA_SHEET_NAME As String = "Data"
Private Const HEADER_ROW As Long = 2
Private Const START_ROW As Long = 3

' Populates data from a recordset into the "Data" sheet.
' @param rs - ADODB.Recordset containing the data
Public Sub PopulateData(rs As ADODB.Recordset)
    Dim targetSheet As Worksheet
    Dim tbl As ListObject
    Dim tblRange As Range
    Dim configManager As ConfigManager
    Dim dbManager As DatabaseManager ' Define DatabaseManager instance
    Dim lastRow As Long, lastCol As Long
    Dim tableName As String

    ' Initialize DatabaseManager
    Set dbManager = New DatabaseManager
    
    ' Initialize ConfigManager
    Set configManager = New ConfigManager
    configManager.Initialize tableName, "", dbManager ' Pass empty SQL string to let configmanager build dynamically

    ' Get the target sheet
    Set targetSheet = ThisWorkbook.Sheets(DATA_SHEET_NAME)

    ' Clear existing data but retain headers
    Call ClearData(targetSheet)

    ' Populate headers and data
    Call PopulateHeaders(targetSheet, rs, configManager)
    Call PopulateRows(targetSheet, rs)

    ' Create table for the populated range
    lastCol = rs.Fields.Count
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Set tblRange = targetSheet.Range(targetSheet.Cells(HEADER_ROW, 1), targetSheet.Cells(lastRow, lastCol))
    Set tbl = CreateDataTable(targetSheet, tblRange)

    ' Apply configurations (validation, tooltips, etc.)
    Call ApplyConfigurations(targetSheet, rs, configManager)
End Sub

' Clears data in the sheet, keeping the headers intact.
Private Sub ClearData(targetSheet As Worksheet)
    targetSheet.Rows(START_ROW & ":" & targetSheet.Rows.Count).ClearContents
End Sub

' Populates headers dynamically based on recordset fields.
Private Sub PopulateHeaders(targetSheet As Worksheet, rs As ADODB.Recordset, configManager As ConfigManager)
    Dim col As Long
    Dim variable As String
    Dim header As String

    For col = 0 To rs.Fields.Count - 1
        variable = rs.Fields(col).Name
        header = configManager.GetColumnMappings()(variable)
        targetSheet.Cells(HEADER_ROW, col + 1).Value = header
    Next col
End Sub

' Populates rows with data from the recordset.
Private Sub PopulateRows(targetSheet As Worksheet, rs As ADODB.Recordset)
    targetSheet.Cells(START_ROW, 1).CopyFromRecordset rs
End Sub

' Converts a range into a formatted table.
Private Function CreateDataTable(targetSheet As Worksheet, tblRange As Range) As ListObject
    Dim tbl As ListObject
    Set tbl = targetSheet.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = "DataTable"
    tbl.TableStyle = "TableStyleMedium9"
    Set CreateDataTable = tbl
End Function

' Applies tooltips and validations using ConfigManager.
Private Sub ApplyConfigurations(targetSheet As Worksheet, rs As ADODB.Recordset, configManager As ConfigManager)
    Dim col As Long
    Dim variable As String

    For col = 0 To rs.Fields.Count - 1
        variable = rs.Fields(col).Name
        Call ApplyTooltip(targetSheet, configManager, variable, HEADER_ROW, col + 1)
        Call ApplyFieldOptions(targetSheet, configManager, variable, START_ROW, col + 1)
        ' Call ApplyDataValidation(targetSheet, configManager, variable, START_ROW, col + 1)
    Next col
End Sub

Private Sub ApplyTooltip(targetSheet As Worksheet, configManager As ConfigManager, variable As String, headerRow As Long, col As Long)
    Dim tooltip As String
    If configManager.GetTooltips.Exists(variable) Then
        tooltip = configManager.GetTooltips()(variable)
        targetSheet.Cells(headerRow, col).AddComment tooltip
    End If
End Sub


' Applies field options for a variable
' @param targetSheet - The Excel sheet to apply the field options to
' @param configManager - The ConfigManager instance
' @param variable - The variable to apply the field options for
' @param startRow - The row number to start applying field options
' @param col - The column number to apply the field options to
Private Sub ApplyFieldOptions(targetSheet As Worksheet, configManager As ConfigManager, variable As String, col As Long)
    Dim fieldOptions As Scripting.Dictionary
    Dim options As String
    
    Set fieldOptions = configManager.GetFieldOptions
    If fieldOptions.Exists(variable) Then
        options = fieldOptions(variable)
        Dim optionArray() As String
        optionArray = Split(options, "|")
        
        Dim optionString As String
        optionString = Join(optionArray, ",")
        
        With targetSheet.Cells(START_ROW, col).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=optionString
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    End If
End Sub

' ' Applies data validation for a variable
' ' @param targetSheet - The Excel sheet to apply the data validation to
' ' @param configManager - The ConfigManager instance
' ' @param variable - The variable to apply the data validation for
' ' @param START_ROW - The row number to start applying data validation
' ' @param col - The column number to apply the data validation to
' Private Sub ApplyDataValidation(targetSheet As Worksheet, configManager As ConfigManager, variable As String, col As Long)
'     Dim dataValidationConfigs As Scripting.Dictionary
'     Dim validation As String
    
'     Set dataValidationConfigs = configManager.GetDataValidationConfigs
'     If dataValidationConfigs.Exists(variable) Then
'         validation = dataValidationConfigs(variable)
'         ' Apply validation rule (example: numeric validation)
'         If validation = "NUMBER" Then
'             With targetSheet.Cells(START_ROW, col).Validation
'                 .Delete
'                 .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:= _
'                 xlBetween, Formula1:="1", Formula2:="100"
'                 .IgnoreBlank = True
'                 .InCellDropdown = True
'                 .ShowInput = True
'                 .ShowError = True
'             End With
'         End If
'     End If
' End Sub