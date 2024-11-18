Option Explicit

Private Const DATA_SHEET_NAME As String = "Data"
Private Const HEADER_ROW As Long = 2
Private Const START_ROW As Long = 3

' Populates data from a recordset into a worksheet.
Public Sub PopulateData(tableName As String)
    Dim dbManager As New DatabaseManager
    Dim configManager As New ConfigManager
    Dim targetSheet As Worksheet
    Dim sqlQuery As String
    Dim rs As ADODB.Recordset

    ' Step 1: Establish database connection
    If Not dbManager.OpenConnection Then
        MsgBox "Failed to connect to the database.", vbCritical
        Exit Sub
    End If

    ' Step 2: Initialize ConfigManager
    configManager.Initialize tableName, dbManager

    ' Step 3: Build data query and fetch data
    sqlQuery = SQLHelper.BuildSelectQuery(tableName)
    Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(sqlQuery, dbManager.GetConnection))

    ' Step 4: Load data into worksheet
    Set targetSheet = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Call ClearData(targetSheet)
    Call PopulateHeaders(targetSheet, rs, configManager)
    Call PopulateRows(targetSheet, rs)

    ' Cleanup
    rs.Close
    dbManager.CloseConnection
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

' Populates data rows from the recordset.
Private Sub PopulateRows(targetSheet As Worksheet, rs As ADODB.Recordset)
    targetSheet.Cells(START_ROW, 1).CopyFromRecordset rs
End Sub
