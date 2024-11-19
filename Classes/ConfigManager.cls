' ConfigManager.cls
Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary
Private dbManager As databaseManager

' Initializes the configuration manager.
' @param tableName - The table name to fetch metadata for.
' @param dbManager - The database manager instance.

Public Sub Initialize(tableName As String, schema As String, dbManager As databaseManager, Optional skipMetadata As Boolean = False)
    Debug.Print "Is ConfigManager instantiated? " & (Not Me Is Nothing)
    
    If dbManager Is Nothing Then
        Err.Raise vbObjectError + 1, "ConfigManager", "DatabaseManager instance is not set."
    End If
    If Not dbManager.IsConnected Then
        Err.Raise vbObjectError + 1, "ConfigManager", "Database connection is not open."
    End If
    

    ' It will skip the load config step if there are mutliple joins
    If Not skipMetadata Then
        Call LoadConfigurationsFromOracle(tableName, schema, dbManager)
    End If
End Sub

' Loads column mappings and data validation configurations from the database.
' @param the databasemanager instance
Private Sub LoadConfigurationsFromOracle(tableName As String, schema As String, dbManager As databaseManager)
    Dim rs As ADODB.Recordset
    Dim metadataquery As String
    Debug.Print "Table Name: " & tableName
    Debug.Print "Schema Name: " & schema
    
    ' Use SQLHelper to build the metadata query
    metadataquery = SQLHelper.BuildMetadataQuery(tableName, schema)

    ' Execute the query and fetch metadata
    Set rs = dbManager.ExecuteQuery(metadataquery)
    
    
    Set columnMappings = New Scripting.Dictionary
    Set dataValidationConfigs = New Scripting.Dictionary

'    Loop through the recordset and populate the dictionaries.
    Do While Not rs.EOF
        columnMappings.Add rs.Fields("COLUMN_NAME").value, rs.Fields("COLUMN_NAME").value
        dataValidationConfigs.Add rs.Fields("COLUMN_NAME").value, rs.Fields("DATA_TYPE").value
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
