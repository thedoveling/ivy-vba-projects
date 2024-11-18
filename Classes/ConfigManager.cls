Option Explicit

Private columnMappings As Object
Private dataValidationConfigs As Object
Private fieldOptions As Object
Private tooltips As Object

' Initializes configurations dynamically from Oracle and the local Config worksheet.
Public Sub Initialize()
    ' Initialize dictionaries
    Set columnMappings = CreateObject("Scripting.Dictionary")
    Set dataValidationConfigs = CreateObject("Scripting.Dictionary")
    Set fieldOptions = CreateObject("Scripting.Dictionary")
    Set tooltips = CreateObject("Scripting.Dictionary")
    
    ' Load configurations
    Call LoadConfigurationsFromOracle("SELECT variable, datatype FROM your_table")
    Call LoadLocalConfigurations
End Sub

' Fetch column mappings dynamically.
Public Function GetColumnMappings() As Object
    Set GetColumnMappings = columnMappings
End Function

' Fetch data validation configurations dynamically.
Public Function GetDataValidationConfigs() As Object
    Set GetDataValidationConfigs = dataValidationConfigs
End Function

' Fetch dropdown field options dynamically.
Public Function GetFieldOptions() As Object
    Set GetFieldOptions = fieldOptions
End Function

' Fetch tooltips dynamically.
Public Function GetTooltips() As Object
    Set GetTooltips = tooltips
End Function

' Load configurations from Oracle.
Private Sub LoadConfigurationsFromOracle(sqlQuery As String)
    Dim dbManager As New DatabaseManager
    Dim rs As ADODB.Recordset
    Dim variable As String, header As String, datatype As String

    On Error GoTo LoadError

    ' Use DatabaseManager to execute the query
    If dbManager.OpenConnectionWithCredentials("password") Then
        Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(sqlQuery, dbManager.GetConnection))
        
        Do While Not rs.EOF
            variable = rs.Fields("variable").Value
            header = variable ' Assuming default header equals variable
            datatype = rs.Fields("datatype").Value
            
            ' Populate dictionaries dynamically
            If Not columnMappings.Exists(variable) Then columnMappings(variable) = header
            If Not dataValidationConfigs.Exists(variable) Then dataValidationConfigs(variable) = datatype
            
            rs.MoveNext
        Loop
        
        rs.Close
        dbManager.CloseConnection
    End If

    Exit Sub

LoadError:
    Call HandleRuntimeError("Error loading configurations from Oracle")
End Sub

' Load local configurations from the Config worksheet.
Private Sub LoadLocalConfigurations()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim variable As String, options As String, tooltip As String

    ' Set the Config worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop through the Config sheet to populate configurations
    For i = 2 To lastRow
        variable = ws.Cells(i, 1).Value
        options = ws.Cells(i, 2).Value
        tooltip = ws.Cells(i, 3).Value

        ' Populate dictionaries dynamically
        If Not fieldOptions.Exists(variable) Then fieldOptions(variable) = options
        If Not tooltips.Exists(variable) Then tooltips(variable) = tooltip
    Next i
End Sub
