' EmailManager.bas
Option Explicit

' Creates an email draft with placeholder replacements.
' @param placeholders - A dictionary of placeholders and their replacements
' @param emailBody - The email body template
' @return - The email body with placeholders replaced
Public Function CreateEmailDraft(placeholders As Scripting.Dictionary, emailBody As String) As String
    Dim placeholder As Variant
    For Each placeholder In placeholders.Keys
        emailBody = Replace(emailBody, placeholder, placeholders(placeholder))
    Next placeholder
    CreateEmailDraft = emailBody
End Function

' Sends an email using Outlook.
' @param toAddress - The recipient's email address
' @param subject - The email subject
' @param body - The email body
Public Sub SendEmail(toAddress As String, subject As String, body As String)
    Dim outlookApp As Object
    Dim mailItem As Object
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)
    
    With mailItem
        .To = toAddress
        .Subject = subject
        .Body = body
        .Send
    End With
End Sub