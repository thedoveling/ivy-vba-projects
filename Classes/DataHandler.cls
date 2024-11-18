Option Explicit

Private Const DATA_SHEET_NAME As String = "Data"
Private Const HEADER_ROW As Long = 2
Private Const START_ROW As Long = 3

' Populates data from a recordset into a worksheet.
Public Sub PopulateData(rs As ADODB.Recordset, Optional useMetadata As Boolean = True)
    Dim dbManager As DatabaseManager
    Dim configManager As ConfigManager
    Dim targetSheet As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim tblRange As Range
    Dim tbl As ListObject

    ' Initialize ConfigManager only if metadata is required
    If useMetadata Then
        Set configManager = New ConfigManager
        configManager.Initialize tableName, dbManager
    End If

    ' Set the target worksheet
    Set targetSheet = ThisWorkbook.Sheets(DATA_SHEET_NAME)

    ' Clear existing data
    Call ClearData(targetSheet)

    ' Populate headers and rows
    If useMetadata Then
        Call PopulateHeaders(targetSheet, rs, configManager)
    Else
        Call PopulateHeadersWithoutMetadata(targetSheet, rs)
    End If

    Call PopulateRows(targetSheet, rs)

    ' Create table and apply configurations (if needed)
    lastCol = rs.Fields.Count
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Set tblRange = targetSheet.Range(targetSheet.Cells(HEADER_ROW, 1), targetSheet.Cells(lastRow, lastCol))
    Set tbl = CreateDataTable(targetSheet, tblRange)

    If useMetadata Then Call ApplyConfigurations(targetSheet, rs, configManager)

    ' Cleanup
    rs.Close
End Sub

' Clears worksheet data while keeping headers intact.
Private Sub ClearData(targetSheet As Worksheet)
    targetSheet.Rows(START_ROW & ":" & targetSheet.Rows.Count).ClearContents
End Sub

' Populates headers dynamically using metadata.
Private Sub PopulateHeaders(targetSheet As Worksheet, rs As ADODB.Recordset, configManager As ConfigManager)
    Dim col As Long
    For col = 0 To rs.Fields.Count - 1
        targetSheet.Cells(HEADER_ROW, col + 1).Value = configManager.GetColumnMappings()(rs.Fields(col).Name)
    Next col
End Sub

Private Sub PopulateHeadersWithoutMetadata(targetSheet As Worksheet, rs As ADODB.Recordset)
    Dim col As Long
    For col = 0 To rs.Fields.Count - 1
        targetSheet.Cells(HEADER_ROW, col + 1).Value = rs.Fields(col).Name
    Next col
End Sub

' Populates data rows from the recordset.
Private Sub PopulateRows(targetSheet As Worksheet, rs As ADODB.Recordset)
    targetSheet.Cells(START_ROW, 1).CopyFromRecordset rs
End Sub

