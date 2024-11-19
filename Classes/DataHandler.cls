' DataHandler.cls
Option Explicit

' Populates data from a recordset into a worksheet.
''' <param name="rs">The ADODB.Recordset containing the data.</param>
''' <param name="useMetadata">Optional. Specifies whether to use metadata for populating headers. Default is True.</param>
Public Sub PopulateData(rs As ADODB.Recordset, useMetadata As Boolean, Optional configManager As ConfigManager = Nothing)
    Dim targetSheet As Worksheet
    Dim lastRow As Long, lastCol As Long, tblRange As Range, tbl As ListObject

    ' Ensure metadata dependencies are provided when required
    If useMetadata And configManager Is Nothing Then
        Err.Raise vbObjectError + 1, "DataHandler", "ConfigManager must be provided when useMetadata is True."
    End If

'   Clear existing data in the target sheet
    Set targetSheet = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    ClearData targetSheet

'    Populate headers based on metadata or without metadata
    If useMetadata Then
        PopulateHeaders targetSheet, rs, configManager
    Else
        PopulateHeadersWithoutMetadata targetSheet, rs
    End If

    PopulateRows targetSheet, rs

'   Create a DataTable from the populated range
    lastCol = rs.Fields.Count
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Set tblRange = targetSheet.Range(targetSheet.Cells(HEADER_ROW, 1), targetSheet.Cells(lastRow, lastCol))
    Set tbl = CreateDataTable(targetSheet, tblRange)
End Sub

''' <summary>
''' Clears the data in the target worksheet.
''' </summary>
''' <param name="targetSheet">The worksheet to clear the data from.</param>
Private Sub ClearData(targetSheet As Worksheet)
    targetSheet.Rows(START_ROW & ":" & targetSheet.Rows.Count).ClearContents
End Sub

''' <summary>
''' Populates the headers in the target worksheet using metadata from the ConfigManager.
''' </summary>
''' <param name="targetSheet">The worksheet to populate the headers in.</param>
''' <param name="rs">The ADODB.Recordset containing the data.</param>
''' <param name="configManager">The ConfigManager object used to retrieve column mappings.</param>
Private Sub PopulateHeaders(targetSheet As Worksheet, rs As ADODB.Recordset, configManager As ConfigManager)
    Dim col As Long
    For col = 0 To rs.Fields.Count - 1
        targetSheet.Cells(HEADER_ROW, col + 1).Value = configManager.GetColumnMappings()(rs.Fields(col).Name)
    Next col
End Sub

''' <summary>
''' Populates the headers in the target worksheet without using metadata.
''' </summary>
''' <param name="targetSheet">The worksheet to populate the headers in.</param>
''' <param name="rs">The ADODB.Recordset containing the data.</param>
Private Sub PopulateHeadersWithoutMetadata(targetSheet As Worksheet, rs As ADODB.Recordset)
    Dim col As Long
    For col = 0 To rs.Fields.Count - 1
        targetSheet.Cells(HEADER_ROW, col + 1).Value = rs.Fields(col).Name
    Next col
End Sub

''' <summary>
''' Populates the rows in the target worksheet with data from the recordset.
''' </summary>
''' <param name="targetSheet">The worksheet to populate the rows in.</param>
''' <param name="rs">The ADODB.Recordset containing the data.</param>
Private Sub PopulateRows(targetSheet As Worksheet, rs As ADODB.Recordset)
    targetSheet.Cells(START_ROW, 1).CopyFromRecordset rs
End Sub

''' <summary>
''' Creates a DataTable from the specified range in the target worksheet.
''' </summary>
''' <param name="targetSheet">The worksheet containing the range.</param>
''' <param name="tblRange">The range to create the DataTable from.</param>
''' <returns>The created ListObject representing the DataTable.</returns>
Private Function CreateDataTable(targetSheet As Worksheet, tblRange As Range) As ListObject
    Dim tbl As ListObject
    Set tbl = targetSheet.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.TableStyle = "TableStyleMedium9"
    Set CreateDataTable = tbl
End Function