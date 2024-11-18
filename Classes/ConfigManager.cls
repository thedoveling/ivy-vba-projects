Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary
Private tooltips As Scripting.Dictionary

' Initializes configurations from Oracle and local "Config" sheet.
Public Sub Initialize()
    Set columnMappings = New Scripting.Dictionary
    Set dataValidationConfigs = New Scripting.Dictionary
    Set tooltips = New Scripting.Dictionary

    ' Load configurations from Oracle
    Call LoadConfigurationsFromOracle("SELECT variable, datatype FROM my_table")

    ' Load local configurations
    Call LoadLocalConfigurations
End Sub

Private Sub LoadConfigurationsFromOracle(sqlQuery As String)
    Dim dbManager As New DatabaseManager
    Dim rs As ADODB.Recordset
    Dim variable As String, header As String, datatype As String

    If dbManager.OpenConnectionWithCredentials("password") Then
        Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(sqlQuery, dbManager.GetConnection))
        Do While Not rs.EOF
            variable = rs.Fields("variable").Value
            header = variable
            datatype = rs.Fields("datatype").Value
            columnMappings(variable) = header
            dataValidationConfigs(variable) = datatype
            rs.MoveNext
        Loop
        rs.Close
        dbManager.CloseConnection
    End If
End Sub

Private Sub LoadLocalConfigurations()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim variable As String, options As String, tooltip As String

    Set ws = ThisWorkbook.Sheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        variable = ws.Cells(i, 1).Value
        options = ws.Cells(i, 2).Value
        tooltip = ws.Cells(i, 3).Value
        fieldOptions(variable) = options
        tooltips(variable) = tooltip
    Next i
End Sub
