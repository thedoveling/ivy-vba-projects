' DataHandler.cls
Option Explicit

' Populates data from a recordset into a worksheet.
''' <param name="rs">The ADODB.Recordset containing the data.</param>
''' <param name="useMetadata">Optional. Specifies whether to use metadata for populating headers. Default is True.</param>
Public Sub PopulateData(rsData As ADODB.Recordset, rsMetadata As ADODB.Recordset, configManager As ConfigManager)
    Dim targetSheet As Worksheet
    Dim lastRow As Long, lastCol As Long, tblRange As Range, tbl As ListObject

   ' Set targetsheet workbook
    Set targetSheet = ThisWorkbook.Sheets(DATA_SHEET_NAME)

    ' Clear existing data in the target sheet
    ClearData targetSheet
    DeleteExistingTables targetSheet

    ' Ensure metadata dependencies are provided when required
    If configManager Is Nothing and Not rsMetadata Is Nothing Then
        Err.Raise vbObjectError + 1, "DataHandler", "ConfigManager must be provided when useMetadata is True."
    End If

'    Populate headers based on metadata or without metadata
    If Not rsMetadata Is Nothing Then
        PopulateHeaders targetSheet, rsMetadata, configManager
        MapAndPopulateRows targetSheet, rsData, configManager
    Else
        PopulateHeadersWithoutMetadata targetSheet, rsData
        PopulateRows targetSheet, rsData
    End If


'   Create a DataTable from the populated range
    lastCol = rsData.Fields.Count
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Set tblRange = targetSheet.Range(targetSheet.Cells(HEADER_ROW, 1), targetSheet.Cells(lastRow, lastCol))
    Set tbl = CreateDataTable(targetSheet, tblRange)
End Sub

''' <summary>
''' Clears the data in the target worksheet.
''' </summary>
''' <param name="targetSheet">The worksheet to clear the data from.</param>
Private Sub ClearData(targetSheet As Worksheet)
    ' Clear all data from the target range, including formatting and comments
    With targetSheet.Rows(START_ROW & ":" & targetSheet.Rows.Count)
        .ClearContents
        .ClearFormats
        .ClearComments
    End With
End Sub

' Deletes any existing tables in the target sheet to prevent overlaps.
Private Sub DeleteExistingTables(targetSheet As Worksheet)
    Dim existingTable As ListObject
    For Each existingTable In targetSheet.ListObjects
        existingTable.Delete
    Next existingTable
End Sub

''' Populates the headers in the target worksheet using metadata from the ConfigManager.
''' <param name="targetSheet">The worksheet to populate the headers in.</param>
''' <param name="rs">The ADODB.Recordset containing the data.</param>
''' <param name="configManager">The ConfigManager instance to use for metadata.</param>
Private Sub PopulateHeaders(targetSheet As Worksheet, rsMetadata As ADODB.Recordset, configManager As ConfigManager)
    Dim col As Long
    Dim headerName As String
    Dim columnMappings As Scripting.Dictionary

    ' Get column mappings from ConfigManager
    Set columnMappings = configManager.GetColumnMappings()

    ' Populate headers based on metadata
    rsMetadata.MoveFirst ' Ensure the recordset starts from the first record
    For col = 0 To rsMetadata.RecordCount - 1
        ' Retrieve the value from COLUMN_NAME field
        headerName = columnMappings(rsMetadata.Fields("COLUMN_NAME").Value)
        targetSheet.Cells(HEADER_ROW, col + 1).Value = headerName
        Debug.Print "Header Set: " & headerName
        rsMetadata.MoveNext
    Next col
End Sub


''' <summary>

''' Populates the headers in the target worksheet without using metadata.
''' </summary>
''' <param name="targetSheet">The worksheet to populate the headers in.</param>
''' <param name="rs">The ADODB.Recordset containing the data.</param>
Private Sub PopulateHeadersWithoutMetadata(targetSheet As Worksheet, rsData As ADODB.Recordset)
    Dim col As Long
    For col = 0 To rsData.Fields.Count - 1
        targetSheet.Cells(HEADER_ROW, col + 1).Value = rsData.Fields(col).Name
    Next col
End Sub


' Maps and populates rows in the target worksheet based on headers.
Private Sub MapAndPopulateRows(targetSheet As Worksheet, rsData As ADODB.Recordset, configManager As ConfigManager)
    Dim col As Long
    Dim columnMappings As Scripting.Dictionary
    Dim fieldName As String

    ' Get column mappings from ConfigManager
    Set columnMappings = configManager.GetColumnMappings()

    ' Copy data into rows, mapped to column headers
    Do While Not rsData.EOF
        For col = 0 To rsData.Fields.Count - 1
            fieldName = rsData.Fields(col).Name
            If columnMappings.Exists(fieldName) Then
                targetSheet.Cells(START_ROW + rsData.AbsolutePosition, col + 1).Value = rsData.Fields(col).Value
            Else
                Debug.Print "Field not mapped: " & fieldName
            End If
        Next col
        rsData.MoveNext
    Loop
End Sub


''' <summary>
''' Populates the rows in the target worksheet with data from the recordset.
''' </summary>
''' <param name="targetSheet">The worksheet to populate the rows in.</param>
''' <param name="rs">The ADODB.Recordset containing the data.</param>
Private Sub PopulateRows(targetSheet As Worksheet, rsData As ADODB.Recordset)
    targetSheet.Cells(START_ROW, 1).CopyFromRecordset rsData
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