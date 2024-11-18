' ConfigManager.cls
Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary
Private fieldOptions As Scripting.Dictionary
Private tooltips As Scripting.Dictionary

Private con As ADODB.Connection

' Initializes the ConfigManager by loading configurations from Oracle and local configuration file.
' @param tableName - The name of the table to load configurations for
' @param sqlQuery - The SQL query to execute for loading configurations
Public Sub Initialize(tableName As String, sqlQuery As String)
    Set columnMappings = New Scripting.Dictionary
    Set dataValidationConfigs = New Scripting.Dictionary
    Set fieldOptions = New Scripting.Dictionary
    Set tooltips = New Scripting.Dictionary
    
    ' Open database connection
    Set con = New ADODB.Connection
    con.ConnectionString = "Provider=ORAOLEDB.ORACLE;Data Source=DEVVV;User Id=" & Environ$("UserName") & ";Password=" & "password" & ";"
    con.Open
    
    ' Load configurations from Oracle
    Call LoadConfigurationsFromOracle(tableName, sqlQuery)
    
    ' Load field options and tooltips from local configuration file
    Call LoadLocalConfigurations
    
    ' Close database connection
    con.Close
    Set con = Nothing
End Sub

' Maps column headers to Oracle fields.
' @return - A dictionary of column mappings
Public Function GetColumnMappings() As Scripting.Dictionary
    Set GetColumnMappings = columnMappings
End Function

' Retrieves data validation configurations.
' @return - A dictionary of data validation configurations
Public Function GetDataValidationConfigs() As Scripting.Dictionary
    Set GetDataValidationConfigs = dataValidationConfigs
End Function

' Retrieves field options for drop-downs.
' @return - A dictionary of field options
Public Function GetFieldOptions() As Scripting.Dictionary
    Set GetFieldOptions = fieldOptions
End Function

' Retrieves tooltips for headers.
' @return - A dictionary of tooltips
Public Function GetTooltips() As Scripting.Dictionary
    Set GetTooltips = tooltips
End Function

' Loads configurations from Oracle.
' @param tableName - The name of the table to load configurations for
' @param sqlQuery - The SQL query to execute for loading configurations
Private Sub LoadConfigurationsFromOracle(tableName As String, sqlQuery As String)
    Dim rs As ADODB.Recordset
    Dim variable As String
    Dim header As String
    Dim datatype As String
    
    On Error GoTo LoadError
    
    ' Execute the provided SQL query to get configurations from Oracle
    Set rs = New ADODB.Recordset
    rs.Open sqlQuery, con, adOpenStatic, adLockReadOnly
    
    ' Load configurations from the recordset
    Do While Not rs.EOF
        variable = rs.Fields("variable").Value
        header = variable ' Assuming header is the same as variable
        datatype = rs.Fields("datatype").Value
        
        ' Populate dictionaries dynamically
        If Not columnMappings.Exists(variable) Then columnMappings(variable) = header
        If Not dataValidationConfigs.Exists(variable) Then dataValidationConfigs(variable) = datatype
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub

LoadError:
    Call HandleRuntimeError("Error loading configurations from Oracle")
End Sub

' Load local configurations from the Config worksheet.
Private Sub LoadLocalConfigurations()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim variable As String, options As String, tooltip As String
    Dim headerRow As Range
    Dim colVariable As Long, colOptions As Long, colTooltip As Long

    ' Set the Config worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find the header row
    Set headerRow = ws.ListObjects("Config").HeaderRowRange

    ' Map the columns dynamically
    colVariable = Application.Match("variable", headerRow, 0)
    colOptions = Application.Match("options", headerRow, 0)
    colTooltip = Application.Match("tooltip", headerRow, 0)

    ' Loop through the Config sheet to populate configurations
    For i = 2 To lastRow
        variable = ws.Cells(i, colVariable).Value
        options = ws.Cells(i, colOptions).Value
        tooltip = ws.Cells(i, colTooltip).Value

        ' Populate dictionaries dynamically
        If Not fieldOptions.Exists(variable) Then fieldOptions(variable) = options
        If Not tooltips.Exists(variable) Then tooltips(variable) = tooltip
    Next i
End Sub