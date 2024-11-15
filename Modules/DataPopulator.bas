' DataPopulator.bas
' Module to handle data population in the workbook
Option Explicit

' Populates headers dynamically in the designated header row
' based on field names from a database query or configuration
Public Sub PopulateHeaders(fieldNames As Collection)
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    
    Dim col As Long
    col = 1 ' Start at the first column
    
    ' Clear existing headers in the HEADER_ROW
    wsData.Rows(HEADER_ROW).ClearContents
    
    ' Populate headers based on provided field names
    Dim fieldName As Variant
    For Each fieldName In fieldNames
        wsData.Cells(HEADER_ROW, col).value = fieldName
        col = col + 1
    Next fieldName
End Sub

' Loads data from a recordset into the worksheet starting from the specified row
' Uses CopyFromRecordset for efficient data transfer
Public Sub LoadDataFromRecordset(rs As ADODB.Recordset)
    On Error GoTo LoadDataError
    
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    
    ' Clear previous data in the data range (starting from START_DATA_ROW)
    ClearDataRange wsData
    
    ' Load data from recordset starting at cell A5 (START_DATA_ROW)
    wsData.Cells(START_DATA_ROW, 1).CopyFromRecordset rs
    
    Exit Sub

LoadDataError:
    MsgBox "Error loading data: " & Err.Description, vbCritical
End Sub

' Clears data in the worksheet without affecting headers
' This function clears rows from START_DATA_ROW onward in DATA_SHEET_NAME
Public Sub ClearDataRange(ws As Worksheet)
    On Error Resume Next
    ' Define the range to clear based on the data starting row and last row with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    If lastRow >= START_DATA_ROW Then
        ws.Range(ws.Cells(START_DATA_ROW, 1), ws.Cells(lastRow, ws.Columns.Count)).ClearContents
    End If
    On Error GoTo 0
End Sub
