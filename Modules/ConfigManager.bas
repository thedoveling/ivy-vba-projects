' ConfigManager.bas
' Manages configuration settings, column mappings, and data validation for the workbook
Option Explicit

' Public constants for configuration settings
Public Const HEADER_ROW As Long = 4
Public Const START_DATA_ROW As Long = 5
Public Const DATA_SHEET_NAME As String = "Data"
Public Const CONFIG_SHEET_NAME As String = "Wrap Up Codes"

' Retrieves column mappings from the "Wrap Up Codes" configuration sheet
' Consolidates mapping functions for flexibility and reduced redundancy
' @return - Dictionary of column mappings (Excel header to Oracle field name)
Public Function GetColumnMappings() As Scripting.Dictionary
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    
    Dim mappings As New Scripting.Dictionary
    Dim lastRow As Long, row As Long
    Dim excelHeader As String, oracleField As String
    
    ' Find the last row in the configuration range
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).row
    
    ' Loop through each row to populate the mappings dictionary
    For row = 2 To lastRow ' Assume row 1 is the header row
        excelHeader = wsConfig.Cells(row, 1).value ' Excel header
        oracleField = wsConfig.Cells(row, 2).value ' Oracle field
        
        ' Only add to dictionary if both Excel header and Oracle field are defined
        If excelHeader <> "" And oracleField <> "" Then
            mappings.Add excelHeader, oracleField
        End If
    Next row
    
    Set GetColumnMappings = mappings
End Function

' Retrieves validation lists for specified fields from the configuration sheet
' @return - Dictionary of validation lists (field name to list range)
Public Function GetValidationLists() As Scripting.Dictionary
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    
    Dim validationLists As New Scripting.Dictionary
    Dim lastRow As Long, row As Long
    Dim fieldName As String, validationRange As String
    
    ' Find the last row in the configuration range
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).row
    
    ' Loop through each row to populate the validation dictionary
    For row = 2 To lastRow ' Assume row 1 is the header row
        fieldName = wsConfig.Cells(row, 1).value ' Field name
        validationRange = wsConfig.Cells(row, 2).value ' Validation list range
        
        ' Only add to dictionary if both field name and validation range are defined
        If fieldName <> "" And validationRange <> "" Then
            validationLists.Add fieldName, validationRange
        End If
    Next row
    
    Set GetValidationLists = validationLists
End Function

' Retrieves the column index for a given field name in the header row
' @return - Column index or 0 if not found
Public Function GetColumnIndex(fieldName As String, Optional ws As Worksheet) As Long
    Dim headerCell As Range
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    
    On Error Resume Next
    Set headerCell = ws.Rows(HEADER_ROW).Find(What:=fieldName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    
    If Not headerCell Is Nothing Then
        GetColumnIndex = headerCell.column
    Else
        GetColumnIndex = 0 ' Returns 0 if header not found
    End If
End Function

' Retrieves a range for a named validation list from CONFIG_SHEET_NAME
' @param validationName - Name of the validation list to retrieve
' @return - Range object for the validation list or Nothing if not found
Public Function GetValidationRange(validationName As String) As Range
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    
    On Error Resume Next
    Set GetValidationRange = wsConfig.Range(validationName)
    On Error GoTo 0
End Function

' Applies data validation to specified columns based on the config
' Validates each configuration range before applying
' @param ws - The worksheet where validation will be applied
' @param validationLists - Dictionary of validation lists (field name to list range)
Public Sub ApplyDataValidations(ByRef ws As Worksheet, ByRef validationLists As Scripting.Dictionary)
    Dim fieldName As Variant
    Dim validationRange As String
    Dim targetRange As Range
    Dim column As Long
    Dim lastRow As Long
    
    ' Get the last row of data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' Loop through each field and apply validation
    For Each fieldName In validationLists.Keys
        column = Application.Match(fieldName, ws.Rows(HEADER_ROW), 0)
        
        If Not IsError(column) Then
            validationRange = validationLists(fieldName)
            
            ' Validate and define the range where validation should be applied
            Set targetRange = ws.Range(ws.Cells(START_DATA_ROW, column), ws.Cells(lastRow, column))
            If ws.Parent.Names(validationRange) Is Nothing Then
                Debug.Print "Warning: Validation range '" & validationRange & "' not found."
            Else
                With targetRange.Validation
                    .Delete
                    .Add type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                        xlBetween, Formula1:="='" & CONFIG_SHEET_NAME & "'!" & validationRange
                End With
            End If
        End If
    Next fieldName
End Sub

' Example method to load mappings and validation configurations
Public Sub LoadMappingsAndValidations()
    Dim mappings As Dictionary
    Dim validationLists As Dictionary
    
    ' Retrieve mappings and validation lists
    Set mappings = GetColumnMappings()
    Set validationLists = GetValidationLists()
    
    ' Log output (optional)
    Dim field As Variant
    For Each field In mappings.Keys
        Debug.Print "Field: " & field & ", Oracle Field: " & mappings(field)
    Next field
    
    For Each field In validationLists.Keys
        Debug.Print "Validation Field: " & field & ", Range: " & validationLists(field)
    Next field
End Sub

' Retrieve configuration dictionary for a specific communication type
' @param configType - Type of communication: "Email" or "Word"
' @return - Dictionary of configuration settings
Public Function GetCommunicationConfig(configType As String) As Scripting.Dictionary
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    
    Dim configs As New Scripting.Dictionary
    Dim lastRow As Long, row As Long
    
    ' Find the last row in the configuration range
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).row
    
    ' Loop through each row to retrieve settings for the specified communication type
    For row = 2 To lastRow
        If wsConfig.Cells(row, 1).value = configType Then
            configs.Add wsConfig.Cells(row, 2).value, wsConfig.Cells(row, 3).value
        End If
    Next row
    
    Set GetCommunicationConfig = configs
End Function
