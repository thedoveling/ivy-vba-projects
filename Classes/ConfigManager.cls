' ConfigManager.cls
Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary
Private dbManager As DatabaseManager

' Initializes the configuration manager.
' @param tableName - The table name to fetch metadata for.
' @param dbManager - The database manager instance.

Public Sub Initialize(tableName As String, dbManager As DatabaseManager, Optional skipMetadata As Boolean = False)
        ' Validate DatabaseManager instance
    If dbManager Is Nothing Then
        Err.Raise vbObjectError + 1, "ConfigManager", "DatabaseManager instance is not set."
    End If
    If Not dbManager.IsConnected Then 
        Err.Raise vbObjectError + 1, "ConfigManager", "Database connection is not open."
    End If

    
    If Not skipMetadata Then 
        Call LoadConfigurationsFromOracle (tableName, dbManager)
    End If
End Sub

' Loads column mappings and data validation configurations from the database.
Private Sub LoadConfigurationsFromOracle(tableName As StringdbManager As DatabaseManager, )
    Dim rs As ADODB.Recordset
    ' NEed to define metadataquery
    Set rs = dbManager.ExecuteQuery(metadataquery)

    Set columnMappings = New Scripting.Dictionary
    Set dataValidationConfigs = New Scripting.Dictionary

'    Loop through the recordset and populate the dictionaries.
    Do While Not rs.EOF
        columnMappings.Add rs.Fields("COLUMN_NAME").Value, rs.Fields("COLUMN_NAME").Value
        dataValidationConfigs.Add rs.Fields("COLUMN_NAME").Value, rs.Fields("DATA_TYPE").Value
        rs.MoveNext
    Loop
    rs.Close
End Sub

' Retrieves the column mappings.
Public Function GetColumnMappings() As Scripting.Dictionary
    Set GetColumnMappings = columnMappings
End Function

' Retrieves the data validation configurations.
Public Function GetDataValidationConfigs() As Scripting.Dictionary
    Set GetDataValidationConfigs = dataValidationConfigs
End Function
