' HeaderManager.bas
Option Explicit

' Set headers dynamically in the HEADER_ROW based on a list of headers
' headers: Array of header names to populate in row HEADER_ROW
Public Sub SetHeaders(headers As Variant)
    Dim ws As Worksheet
    Dim i As Integer

    ' Get the data sheet
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)

    ' Loop through headers and set each in HEADER_ROW
    For i = LBound(headers) To UBound(headers)
        ws.Cells(HEADER_ROW, i + 1).value = headers(i)
    Next i
End Sub

' Add tooltips to each header cell in the HEADER_ROW based on configuration in TOOLTIP_SHEET_NAME
Public Sub AddTooltips()
    Dim wsData As Worksheet
    Dim wsConfig As Worksheet
    Dim headers As Range
    Dim tooltipRange As Range
    Dim headerCell As Range
    Dim tooltipCell As Range
    Dim headerName As String
    Dim tooltipText As String
    
    ' Set references to data and tooltip sheets
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    Set wsConfig = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    
    ' Assume headers are in row HEADER_ROW of the data sheet
    Set headers = wsData.Rows(HEADER_ROW)
    
    ' Assume tooltips are in columns A (header names) and B (tooltips) in the tooltip sheet
    Set tooltipRange = wsConfig.Range("A1", wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp))
    
    ' Loop through each header in the HEADER_ROW
    For Each headerCell In headers
        headerName = headerCell.value
        If Not IsEmpty(headerName) Then
            ' Find the tooltip for this header
            Set tooltipCell = tooltipRange.Find(headerName, LookIn:=xlValues, LookAt:=xlWhole)
            If Not tooltipCell Is Nothing Then
                tooltipText = tooltipCell.Offset(0, 1).value
                ' Add tooltip to the header cell
                If Len(tooltipText) > 0 Then
                    ' Remove existing comment before adding a new one
                    If Not headerCell.Comment Is Nothing Then headerCell.Comment.Delete
                    headerCell.AddComment tooltipText
                End If
            End If
        End If
    Next headerCell
End Sub
