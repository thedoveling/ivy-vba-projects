Public Sub ApplyDataValidation()

    Dim categoryCol As Long, outcomeCol As Long, subReasonCol As Long
    Dim sendEmailCol As Long, slAddCol As Long, armIssueCol As Long, reqResponseCol As Long
    Dim categoryRange As Range, outcomeRange As Range, subReasonRange As Range
    Dim sendEmailRange As Range, sampleLossRange As Range, armIssueRange As Range, reqResponseRange As Range
    Dim categoryList As Range, outcomeList As Range, subReasonList As Range, ynFlagList As Range

    Debug.Print "Headers Address: " & headers.address
    
    ' Find last row of data in the Data worksheet
    lastrow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row

    ' Find the columns of interest in the Data worksheet
    categoryCol = Application.Match("Category", headers, 0)
    outcomeCol = Application.Match("Outcome", headers, 0)
    subReasonCol = Application.Match("Sub_Reason", headers, 0)
    sendEmailCol = Application.Match("SEND_EMAIL", headers, 0)
    slAddCol = Application.Match("SL_ADD", headers, 0)
    armIssueCol = Application.Match("ARM_ISSUE", headers, 0)
    reqResponseCol = Application.Match("REQUESTED_RESPONSE", headers, 0)

    ' Set the list ranges in the Wrap Up Codes worksheet
    Set categoryList = wsWrapUp.ListObjects("Category").ListColumns(1).DataBodyRange
    Set outcomeList = wsWrapUp.ListObjects("Outcome").ListColumns(1).DataBodyRange
    Set subReasonList = wsWrapUp.ListObjects("Sub_reason").ListColumns(1).DataBodyRange
    Set ynFlagList = wsWrapUp.ListObjects("YN_FLAG").ListColumns(1).DataBodyRange

    ' Apply data validation for "Category" using the range address of the Category table in Wrap Up Codes
    If Not IsError(categoryCol) Then
        Set categoryRange = wsData.Range(wsData.Cells(5, categoryCol), wsData.Cells(lastrow, categoryCol))
        categoryRange.Validation.Delete
        categoryRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                    xlBetween, Formula1:="='" & wsWrapUp.Name & "'!" & categoryList.address
    End If

    ' Apply data validation for "Outcome" using the range address of the Outcome table in Wrap Up Codes
    If Not IsError(outcomeCol) Then
        Set outcomeRange = wsData.Range(wsData.Cells(5, outcomeCol), wsData.Cells(lastrow, outcomeCol))
        outcomeRange.Validation.Delete
        outcomeRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                    xlBetween, Formula1:="='" & wsWrapUp.Name & "'!" & outcomeList.address
    End If

    ' Apply data validation for "Sub_Reason" using the range address of the Sub_Reason table in Wrap Up Codes
    If Not IsError(subReasonCol) Then
        Set subReasonRange = wsData.Range(wsData.Cells(5, subReasonCol), wsData.Cells(lastrow, subReasonCol))
        subReasonRange.Validation.Delete
        subReasonRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                      xlBetween, Formula1:="='" & wsWrapUp.Name & "'!" & subReasonList.address
    End If

    ' Apply Y/N validation for "SEND_EMAIL" using the range address of the YN_FLAG table in Wrap Up Codes
    If Not IsError(sendEmailCol) Then
        Set sendEmailRange = wsData.Range(wsData.Cells(5, sendEmailCol), wsData.Cells(lastrow, sendEmailCol))
        sendEmailRange.Validation.Delete
        sendEmailRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                      xlBetween, Formula1:="='" & wsWrapUp.Name & "'!" & ynFlagList.address
    End If

    ' Apply Y/N validation for "SL_ADD" using the range address of the YN_FLAG table in Wrap Up Codes
    If Not IsError(slAddCol) Then
        Set sampleLossRange = wsData.Range(wsData.Cells(5, slAddCol), wsData.Cells(lastrow, slAddCol))
        sampleLossRange.Validation.Delete
        sampleLossRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                       xlBetween, Formula1:="='" & wsWrapUp.Name & "'!" & ynFlagList.address
    End If

    ' Apply Y/N validation for "ARM_ISSUE" using the range address of the YN_FLAG table in Wrap Up Codes
    If Not IsError(armIssueCol) Then
        Set armIssueRange = wsData.Range(wsData.Cells(5, armIssueCol), wsData.Cells(lastrow, armIssueCol))
        armIssueRange.Validation.Delete
        armIssueRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                     xlBetween, Formula1:="='" & wsWrapUp.Name & "'!" & ynFlagList.address
    End If

    ' Apply Y/N validation for "REQUESTED_RESPONSE" using the range address of the YN_FLAG table in Wrap Up Codes
    If Not IsError(reqResponseCol) Then
        Set reqResponseRange = wsData.Range(wsData.Cells(5, reqResponseCol), wsData.Cells(lastrow, reqResponseCol))
        reqResponseRange.Validation.Delete
        reqResponseRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                        xlBetween, Formula1:="='" & wsWrapUp.Name & "'!" & ynFlagList.address
    End If

End Sub


Public Sub AddHeaderTooltips()
    Dim tooltipCol As String
    Dim commentText As String
    
    Debug.Print "Headers Address: " & headers.address

    ' Loop through each header globally defined in headers
    For Each header In headers
        ' Make sure the header contains a value and is not an error
        If Not IsEmpty(header.Value) And Not IsError(header.Value) Then
            tooltipCol = CStr(header.Value) ' Convert to string to avoid type mismatch errors
            
            Debug.Print "Header Value: " & tooltipCol

            ' Reset commentText for each header to avoid retaining old values
            commentText = ""

            ' Set the tooltip text based on the header value
            Select Case UCase(tooltipCol) ' Match against uppercased header names
                Case "ID"
                    commentText = "The ID in the sweep tool."
                Case "DATE_RECEIVED"
                    commentText = "Date the query was received by Sampling. Not the same as the date the query was submitted."
                Case "SURVEY"
                    commentText = "The survey associated with the query"
                Case "ADDRESS"
                    commentText = "The physical address associated with the Master Sample ID."
                Case "MASTER_SAMPLE_ID"
                    commentText = "This is the Master Sample ID which uniquely identifies a record. Otherwise known as ProviderID."
                Case "HOUSEHOLD_ID"
                    commentText = "This is the obligation ID from Pega or Indicative from CAIWMS."
                Case "INTERVIEWER"
                    commentText = "The Field Interviewer currently assigned to the record."
                Case "INTERACTION_COMMENTS"
                    commentText = "The remarks associated with the query."
                Case "CATEGORY"
                    commentText = "The category used to submit the query to Sampling."
                Case "INVESTIGATION"
                    commentText = "Max character length is 2000. Include the full investigation techniques and outcomes. This is for internal use."
                Case "OUTCOME"
                    commentText = "Outcome of the investigation. Record the final result here using the drop-down list."
                Case "SUB_REASON"
                    commentText = "Here we record the main reason associated with the outcome. This will inform the sample loss tag and email to Field Operations."
                Case "OMT_RESPONSE"
                    commentText = "Our final investigation outcome for Field Operations use. This will populate in your email draft and is visible in the BAT."
                Case "SEND_EMAIL"
                    commentText = "Use this as a switch for the draft email button. An email draft will only generate if this is Y."
                Case "SL_ADD"
                    commentText = "Set this to Y to indicate whether the sample should be flagged as a loss in the Lunar system. The Sample Loss button will only create a tag if this = Y."
                Case "ARM_ISSUE"
                    commentText = "This denotes whether there is an address issue to be rectified in the Address Register."
                Case "REQUESTED_RESPONSE"
                    commentText = "Set this to Y if you require a response from the Field Interviewer. Otherwise set to N."
                Case "CHECKED_BY"
                    commentText = "Enter user ID to assign work."
                Case "QSAMPLELOSS"
                    commentText = "Read only. This is the sample loss variable in key Lunar sample management table zzzhsf.Lunar_dwellotherinfo."
                Case "QOTHERIDENTIFICATION"
                    commentText = "Read only. This is the Other ID variable in in key Lunar sample management table zzzhsf.Lunar_dwellotherinfo. This should generally align with the txOtherIdentification field in CAIWMS."
                Case "QOHASCOMMENT"
                    commentText = "Read only. This is the OHAS comment variable in in key Lunar sample management table zzzhsf.Lunar_dwellotherinfo.  This should generally align with the txOHASConcerns field in CAIWMS."
                Case "QOHASCATEGORY"
                    commentText = "Read only. This is the OHAS category variable in in key Lunar sample management table zzzhsf.Lunar_dwellotherinfo. Code field labels available above."
                Case "QNOCAPI"
                    commentText = "Read only. This is the do not approach OHAS variable in in key Lunar sample management table zzzhsf.Lunar_dwellotherinfo. This will be set to Y to indicate that an address should not be visited in person."
                Case "DATE_COMPLETED"
                    commentText = "Enter todays date to close off the query and signify that the investigation is complete. Work will dissapear from view when this is populated."
                Case "TIME_TO_COMPLETE"
                    commentText = "Enter the approximate time taken to excecute the investigation."
                Case "FURTHER_DETAILS_FOR_ARM"
                    commentText = "Enter any feedback you wish to provide to the Address Register Maintenance team (ARM)."
                Case "IMPROVE_ADDRESS_INFO_USED"
                    commentText = "Set this to Y or N to indicate whether you have used the feedback button directly in the AR user interface."
                Case "PDF_CREATED"
                    commentText = "This is for Lot investigations. Set this to Y when you have generated a PDF."
                Case "STATE"
                    commentText = "Read only. The state associated with the address."
                Case "CHANGED"
                    commentText = "DO NOT UPDATE. This should technically be locked/read only but I need this unlocked to perform important functions. This will set a flag when an update has been made in the sweep tool but not yet committed to Oracle."
                Case "UPDATED"
                    commentText = "DO NOT UPDATE. This should technically be locked/read only but I need this unlocked to perform important functions. This will set to Updated when you have committed changes to Oracle."
                Case "INSERTED"
                    commentText = "DO NOT UPDATE. This should technically be locked/read only but I need this unlocked to perform important functions."
                Case "ARID"
                    commentText = "Use this to quickly view the address in the AR UI. Do not reference this ID in your investigation or save this anywhere."
                Case "COUNT"
                    commentText = "If count is more than one, that means we've seen the master_sample_ID before in our sweep tool. Investigate."

                ' Add more cases for other headers as needed
                Case Else
                    commentText = "No information available."
            End Select

            ' Add the comment to the header cell if there's a comment text to add
            If Len(commentText) > 0 Then
                ' Remove any existing comment before adding a new one
                If Not header.Comment Is Nothing Then header.Comment.Delete
                header.AddComment commentText
            End If
        End If
    Next header
End Sub



