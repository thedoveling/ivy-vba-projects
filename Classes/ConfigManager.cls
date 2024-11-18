Option Explicit

Private columnMappings As Scripting.Dictionary
Private dataValidationConfigs As Scripting.Dictionary
Private dbManager As DatabaseManager

' Initializes the configuration manager.
' @param tableName - The table name to fetch metadata for.
' @param dbManager - The database manager instance.
Public Sub Initialize(tableName As String, dbManager As DatabaseManager)
    If Not dbManager.IsConnected Then
        Err.Raise vbObjectError + 1, "ConfigManager", "Database connection is not open."
    End If

    ' Set DatabaseManager instance
    Set Me.dbManager = dbManager

    ' Load metadata configurations
    Call LoadConfigurationsFromOracle(tableName)
End Sub

' Loads metadata configurations from Oracle.
Private Sub LoadConfigurationsFromOracle(tableName As String)
    Dim rs As ADODB.Recordset
    Dim metadataQuery As String

        ' Use SQLHelper to build the metadata query
    metadataQuery = SQLHelper.BuildMetadataQuery(tableName)
    Set rs = dbManager.ExecuteCommandQuery(dbManager.CreateCommand(metadataQuery, dbManager.GetConnection))

    ' Initialize dictionaries
    Set columnMappings = New Scripting.Dictionary
    Set dataValidationConfigs = New Scripting.Dictionary

    ' Populate metadata
    Do While Not rs.EOF
        columnMappings.Add rs.Fields("COLUMN_NAME").Value, rs.Fields("COLUMN_NAME").Value
        dataValidationConfigs.Add rs.Fields("COLUMN_NAME").Value, rs.Fields("DATA_TYPE").Value
        rs.MoveNext
    Loop
    rs.Close
End Sub

' Provides access to column mappings.
Public Function GetColumnMappings() As Scripting.Dictionary
    Set GetColumnMappings = columnMappings
End Function

' Provides access to data validation configurations.
Public Function GetDataValidationConfigs() As Scripting.Dictionary
    Set GetDataValidationConfigs = dataValidationConfigs
End Function
