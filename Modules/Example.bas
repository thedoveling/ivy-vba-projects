' UnitTesting.bas
Option Explicit

' Runs all tests.
Public Sub RunAllTests()
    On Error GoTo TestError

    Debug.Print "Starting Tests..."
    Call Test_ConfigManager
    Call Test_DataHandler
    Debug.Print "All Tests Completed Successfully."

    Exit Sub

TestError:
    Call HandleRuntimeError("Test failed")
End Sub

' Test ConfigManager initialization.
Private Sub Test_ConfigManager()
    Dim configManager As ConfigManager
    Set configManager = New ConfigManager

    configManager.Initialize
    Debug.Assert Not configManager.GetColumnMappings Is Nothing
    Debug.Assert configManager.GetColumnMappings.Count > 0

    Debug.Print "ConfigManager Tests Passed."
End Sub

' Test DataHandler population.
Private Sub Test_DataHandler()
    Dim dataHandler As DataHandler
    Dim mockRecordset As ADODB.Recordset
    Set dataHandler = New DataHandler

    ' Create mock recordset
    Set mockRecordset = CreateObject("ADODB.Recordset")
    With mockRecordset
        .Fields.Append "Column1", adVarChar, 50
        .Fields.Append "Column2", adVarChar, 50
        .Open
        .AddNew Array("Column1", "Column2"), Array("Data1", "Data2")
    End With

    ' Call PopulateData
    Call dataHandler.PopulateData(mockRecordset)

    ' Assert
    Debug.Assert ThisWorkbook.Sheets("Data").ListObjects.Count > 0

    Debug.Print "DataHandler Tests Passed."
End Sub
