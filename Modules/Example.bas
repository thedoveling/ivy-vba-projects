' UnitTest_ConfigManager.bas
Option Explicit

Public Sub Test_ConfigManager()
    Dim configManager As ConfigManager
    Dim columnMappings As Object, dataValidationConfigs As Object
    Dim fieldOptions As Object, tooltips As Object

    ' Initialize ConfigManager
    Set configManager = New ConfigManager
    configManager.Initialize

    ' Fetch dynamically loaded data
    Set columnMappings = configManager.GetColumnMappings()
    Set dataValidationConfigs = configManager.GetDataValidationConfigs()
    Set fieldOptions = configManager.GetFieldOptions()
    Set tooltips = configManager.GetTooltips()

    ' Assertions (mock data checks)
    Debug.Assert Not columnMappings Is Nothing
    Debug.Assert Not dataValidationConfigs Is Nothing
    Debug.Assert Not fieldOptions Is Nothing
    Debug.Assert Not tooltips Is Nothing

    ' Check for specific mock values
    Debug.Assert columnMappings.Exists("ID") ' Check Oracle-loaded variable
    Debug.Assert fieldOptions.Exists("ID")    ' Check local Config variable
    Debug.Print "Test_ConfigManager passed."
End Sub
