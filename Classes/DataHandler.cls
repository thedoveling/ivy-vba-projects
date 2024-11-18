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
    Dim lastRow As Long, lastCol As Long

    ' Initialize ConfigManager
    Set configManager = New ConfigManager
    configManager.Initialize

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
        Call ApplyDataValidation(targetSheet, configManager, variable, START_ROW, col + 1)
    Next col
End Sub

Private Sub ApplyTooltip(targetSheet As Worksheet, configManager As ConfigManager, variable As String, headerRow As Long, col As Long)
    Dim tooltip As String
    If configManager.GetTooltips.Exists(variable) Then
        tooltip = configManager.GetTooltips()(variable)
        targetSheet.Cells(headerRow, col).AddComment tooltip
    End If
End Sub

Private Sub ApplyFieldOptions(targetSheet As Worksheet, configManager As ConfigManager, variable As String, startRow As Long, col As Long)
    ' Implementation similar to the original
End Sub

Private Sub ApplyDataValidation(targetSheet As Worksheet, configManager As ConfigManager, variable As String, startRow As Long, col As Long)
    ' Implementation similar to the original
End Sub
