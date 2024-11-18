Option Explicit

Public Sub TestDatabaseFlow()
    On Error GoTo ErrorHandler

    ' Declare required objects
    Dim dbManager As DatabaseManager
    Dim configManager As ConfigManager
    Dim dataHandler As DataHandler
    Dim rs As ADODB.Recordset
    Dim tableName As String
    Dim useMetadata As Boolean

    ' Initialize objects
    Set dbManager = New DatabaseManager
    Set configManager = New ConfigManager
    Set dataHandler = New DataHandler

    ' Set table name and metadata flag
    tableName = "YOUR_TABLE_NAME" ' Update with the actual table name
    useMetadata = True ' Set to False to skip metadata processing

    ' Open database connection
    If Not dbManager.OpenConnection Then
        MsgBox "Connection failed. Please log in again.", vbCritical
        Exit Sub
    End If

    ' Initialize ConfigManager if metadata is required
    If useMetadata Then
        configManager.Initialize tableName, dbManager
    End If

    ' Execute data query
    Set rs = dbManager.ExecuteQuery(SQLHelper.BuildSelectQuery(tableName))
    If rs Is Nothing Or rs.EOF Then
        MsgBox "No data retrieved from table: " & tableName, vbInformation
        Exit Sub
    End If

    ' Populate worksheet with data
    dataHandler.PopulateData rs, useMetadata
    MsgBox "Data successfully populated into worksheet.", vbInformation

    ' Cleanup
    rs.Close
    dbManager.CloseConnection
    Exit Sub

ErrorHandler:
    Call HandleRuntimeError("An error occurred during the test flow.")
    If Not dbManager Is Nothing Then dbManager.CloseConnection
End Sub
