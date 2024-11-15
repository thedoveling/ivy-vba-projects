' WordMailMergeManager.bas
Option Explicit

' Performs a mail merge in Word using the specified worksheet data and configuration
' @param ws - Worksheet containing data for placeholders
' @param startRow - Starting row for data in the worksheet
' @param config - Dictionary of configuration settings (TemplatePath, OutputPath, Field Mapping)
Public Sub PerformMailMerge(ByRef ws As Worksheet, ByVal startRow As Long, ByRef config As Scripting.Dictionary)
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim lastRow As Long, currentRow As Long
    Dim fieldName As String, fieldValue As String
    Dim outputPath As String, fileName As String

    ' Initialize Word application and open template document
    Set WordApp = CreateObject("Word.Application")
    Set WordDoc = WordApp.Documents.Open(config("TemplatePath"))
    
    ' Loop through each row of data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For currentRow = startRow To lastRow
        ' Loop through each field mapping
        For Each fieldName In config.Keys
            If Not IsError(Application.Match(fieldName, ws.Rows(HEADER_ROW), 0)) Then
                fieldValue = ws.Cells(currentRow, Application.Match(fieldName, ws.Rows(HEADER_ROW), 0)).value
                WordDoc.Content.Find.Execute FindText:="[" & fieldName & "]", ReplaceWith:=fieldValue, Replace:=2
            End If
        Next fieldName

        ' Save as a new file
        fileName = ws.Cells(currentRow, 1).value & "_MailMerge.docx" ' Adjust naming as needed
        outputPath = config("OutputPath") & "\" & fileName
        WordDoc.SaveAs outputPath
    Next currentRow

    ' Close Word document
    WordDoc.Close False
    WordApp.Quit
End Sub
