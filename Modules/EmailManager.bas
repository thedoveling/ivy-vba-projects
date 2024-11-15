' EmailManager.bas
Option Explicit

' Creates email drafts in Outlook based on the specified worksheet data and configuration
' @param ws - Worksheet containing data for placeholders
' @param startRow - Starting row for data in the worksheet
' @param config - Dictionary of configuration settings (To, CC, Subject Template, Body Template Path)
Public Sub CreateEmailDrafts(ByRef ws As Worksheet, ByVal startRow As Long, ByRef config As Scripting.Dictionary)
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim lastRow As Long, currentRow As Long
    Dim subjectLine As String, bodyText As String, fieldName As String
    Dim placeholderValue As String

    ' Initialize Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Loop through each row to create email drafts
    For currentRow = startRow To lastRow
        ' Create a new mail item from the specified template
        Set MailItem = OutlookApp.CreateItem(0)
        
        ' Populate To, CC, and Subject using external config or defaults in the template
        MailItem.To = config("To")
        MailItem.CC = config("CC")
        subjectLine = config("SubjectTemplate")
        
        ' Replace subject placeholders with worksheet values
        For Each fieldName In config.Keys
            If Not IsError(Application.Match(fieldName, ws.Rows(HEADER_ROW), 0)) Then
                placeholderValue = ws.Cells(currentRow, Application.Match(fieldName, ws.Rows(HEADER_ROW), 0)).value
                subjectLine = Replace(subjectLine, "[" & fieldName & "]", placeholderValue)
            End If
        Next fieldName
        MailItem.Subject = subjectLine
        
        ' Load body from template file
        bodyText = GetTemplateContent(config("BodyTemplatePath"))
        For Each fieldName In config.Keys
            If Not IsError(Application.Match(fieldName, ws.Rows(HEADER_ROW), 0)) Then
                placeholderValue = ws.Cells(currentRow, Application.Match(fieldName, ws.Rows(HEADER_ROW), 0)).value
                bodyText = Replace(bodyText, "[" & fieldName & "]", placeholderValue)
            End If
        Next fieldName
        MailItem.HTMLBody = bodyText

        ' Save the draft
        MailItem.Save
    Next currentRow
End Sub

' Helper function to retrieve template content from a file
' @param filePath - Path to the text/HTML file
' @return - Content of the template as string
Private Function GetTemplateContent(filePath As String) As String
    Dim fileContent As String
    Dim fileNum As Integer
    
    On Error Resume Next
    fileNum = FreeFile
    Open filePath For Input As fileNum
    If Err.Number = 0 Then
        fileContent = Input$(LOF(fileNum), fileNum)
    End If
    Close fileNum
    GetTemplateContent = fileContent
End Function
