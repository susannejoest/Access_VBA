Attribute VB_Name = "mdl_Outlook"
Option Compare Database

    
Public Function fctCreateEmail_Simple(strTo As String, strCC As String, strSubject As String, strBody As String, blnSend As Boolean, Optional strAttachment As String, Optional strPathNameSaveAs As String)

'Microsoft Outlook 16.0 Library must be referenced!!

On Error GoTo Error_fctCreateEmail_Simple

'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS
Dim olApp As Outlook.Application
Dim olNewMail As Outlook.MailItem
Dim intcount As Integer
'Dim fsoFolder As Variant
'If IsNull(blnCreateIfNoAttachment) Then blnCreateIfNoAttachment = False
'Dim blnAttachment As Boolean
'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS
    Set olApp = GetObject(, "Outlook.Application")
    Set olNewMail = olApp.CreateItem(olMailItem)
    
    ' for multiple recipients use a concat string "intern1@test.com; intern2@test.com; intern3@test.com"
    olNewMail.To = strTo
    
    If Not strCC = "" Then olNewMail.CC = strCC
    
    If Not strSubject = "" Then olNewMail.Subject = strSubject
    
    If Not IsNull(strBody) Then olNewMail.HTMLBody = strBody
    
    If strAttachment > "" Then olNewMail.Attachments.Add strAttachment

lblSend:
    If blnSend Then
        olNewMail.Send
    Else
        If Not IsNull(strPathNameSaveAs) > "" Then
            olNewMail.SaveAs strPathNameSaveAs & ".msg", olMSG
        Else
            olNewMail.Save
        End If
    'OlSaveAsType constants: olHTML, olMSG, olRTF, olTemplate, olDoc, olTXT, olVCal, olVCard, olICal, or olMSGUnicode.
    End If
    

    'Debug.Print strTo
    'Debug.Print "NEW MAIL"
lblCleanup:
    Set olNewMail = Nothing

Exit Function

'ERROR'ERROR'ERROR'ERROR'ERROR'ERROR
Error_fctCreateEmail_Simple:

    Select Case ERR.Number
        'Case 287    'User clicked "no" to the Outlook message box asking if it may send email
            'If MsgBox("You have to click yes to the Outlook message box to allow Access to email. Select YES and then click yes to the Outlook message box. Or click NO to abort!", vbYesNo) = vbYes Then
                'Resume
                'Else    'Abort
                'Exit Function
                'GoTo lblCleanup
            'End If
        Case 429    'Get Object failed because Outlook is not open yet. Open new instance!
            Set olApp = CreateObject("Outlook.Application")
            Resume Next
        Case -2147024894
            MsgBox "Attachment not found : '" & strAttachment & "' (resuming next)"
            Resume Next
        Case Else   'untrapped error
            Debug.Print "Error in fctCreateEmail_Simple " '& arrAttachmentPath(intCount) & Err.Number & Err.DESCRIPTION
            MsgBox "Unknown error creating mail!" & ERR.Number & ERR.Description
            Exit Function
            GoTo lblCleanup
    End Select
    
End Function

**********
Declarations for Email variables
        Dim strTo As String
        Dim strCC As String
        Dim strSubject As String
        Dim strBody As String
        Dim blnSend As Boolean
************

Public Function fctCreateEmail_MultipleRecipients(strTo, strToCount As Integer, strCC As Variant, strSubject As String, strBody As String, blnSend As Boolean, Optional strAttachment As String, Optional blnDisplay As Boolean, Optional strVotingOptions)
'Dim strTo(10) As Variant
' strVoting Options "Yes;No"

On Error GoTo ERR

'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS
Dim olApp As Outlook.Application
Dim olNewMail As Outlook.MailItem
Dim olRecipientTo As Outlook.Recipient

Dim intcount As Integer
'Dim fsoFolder As Variant
'If IsNull(blnCreateIfNoAttachment) Then blnCreateIfNoAttachment = False
'Dim blnAttachment As Boolean
'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS
    Set olApp = GetObject(, "Outlook.Application")
    Set olNewMail = olApp.CreateItem(olMailItem)
    
    intcount = 0
    Do Until intcount = strToCount
        intcount = intcount + 1
        Set olRecipientTo = olNewMail.Recipients.Add(strTo(intcount))
        olRecipientTo.Resolve
    Loop
    
     'If myDelegate.Resolved Then
    
    If Not strCC = "" Then olNewMail.CC = strCC
    
    olNewMail.Subject = strSubject
    
    If Not IsNull(strBody) Then olNewMail.HTMLBody = strBody
    
    If strAttachment > "" Then olNewMail.Attachments.Add strAttachment
    If strVotingOptions > "" Then olNewMail.VotingOptions = strVotingOptions
    
    If blnDisplay = True Then olNewMail.Display
    

lblSend:
    If blnSend Then
        olNewMail.Send
    Else: olNewMail.Save
    End If
    
    'Debug.Print strTo
    'Debug.Print "NEW MAIL"
lblCleanup:
Set olNewMail = Nothing

Exit Function

'ERROR'ERROR'ERROR'ERROR'ERROR'ERROR
ERR:

    Select Case ERR.Number
        'Case 287    'User clicked "no" to the Outlook message box asking if it may send email
            'If MsgBox("You have to click yes to the Outlook message box to allow Access to email. Select YES and then click yes to the Outlook message box. Or click NO to abort!", vbYesNo) = vbYes Then
                'Resume
                'Else    'Abort
                'Exit Function
                'GoTo lblCleanup
            'End If
        Case 429    'Get Object failed because Outlook is not open yet. Open new instance!
            Set olApp = CreateObject("Outlook.Application")
            Resume Next
        Case -2147024894
            MsgBox "Attachment not found : '" & strAttachment & "' (resuming next)"
            Resume Next
        Case Else   'untrapped error
            Debug.Print "Error in fctCreateEmail_Simple " '& arrAttachmentPath(intCount) & Err.Number & Err.DESCRIPTION
            MsgBox "Unknown error creating mail!" & ERR.Number & ERR.Description
            Exit Function
            GoTo lblCleanup
    End Select
    
End Function

Function testMultipleRecipientEmail()
    
    Dim strTo(10) As Variant
    Dim strCC As String
    
    strTo(1) = "Email@google.com"
    strTo(2) = "Email2@google.com"
    
    Call fctCreateEmail_MultipleRecipients(strTo, 2, "", "strSubject As String", "strBody As String", False, , True, "Yes;No")

End Function


Public Function TESTfctCheckEmailAD(strTo As String) As Boolean

On Error GoTo Error_fctCheckEmailAD

'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS'DECLARATIONS
Dim olApp As Outlook.Application
Dim olAL As Outlook.AddressList
Dim olAE As Outlook.AddressEntry
Dim olEU As Outlook.ExchangeUser

Set olApp = GetObject(, "Outlook.Application")
Set olAL = olApp.Session.AddressLists("Global Address List")

'oAL.AddressEntries collection, comparing each AddressEntry
Set olAE = olAL.AddressEntries.GetFirst

Set olEU = olAE.GetExchangeUser
Debug.Print olEU.AssistantName

lblCleanup:
'Set olNewMail = Nothing
'Set olRecipient = Nothing

Exit Function

'ERROR'ERROR'ERROR'ERROR'ERROR'ERROR
Error_fctCheckEmailAD:

    Select Case ERR.Number
        'Case 287    'User clicked "no" to the Outlook message box asking if it may send email
            'If MsgBox("You have to click yes to the Outlook message box to allow Access to email. Select YES and then click yes to the Outlook message box. Or click NO to abort!", vbYesNo) = vbYes Then
                'Resume
                'Else    'Abort
                'Exit Function
                'GoTo lblCleanup
            'End If
        Case 429    'Get Object failed because Outlook is not open yet. Open new instance!
            Set olApp = CreateObject("Outlook.Application")
            Resume Next
        Case -2147024894
            MsgBox "Attachment not found : '" & strAttachment & "' (resuming next)"
            Resume Next
        Case Else   'untrapped error
            Debug.Print "Error in fctCreateEmail_Simple " '& arrAttachmentPath(intCount) & Err.Number & Err.DESCRIPTION
            MsgBox "Unknown error creating mail!" & ERR.Number & ERR.Description
            Exit Function
            GoTo lblCleanup
    End Select
    
End Function

Private Sub cmdCreateEmail_TemplateCode()

' LEARF STAR Decision Matrix

    Dim strTo(10) As Variant
    Dim strCC As String
    Dim strSubject As String
    Dim strBody As String
    Dim strAttachment As String
    Dim blnSend As Boolean
    
    blnSend = False
    'strTo = ""
    strCC = ""
    strSubject = "STAR / LEARF approval: " & Str_Description
    strBody = "Please review and approve. " & vbCrLf & "Description: " & Str_Description
    
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim intcount As Integer
    Dim strVotingOptions As String
    strVotingOptions = "Yes;No"
    
    strSQL = "select * from qry_Req_Approvers_ALL"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    intcount = 1
    
    Do Until rst.EOF
        strTo(intcount) = rst!ApprPersonEmail
        intcount = intcount + 1
        rst.MoveNext
    Loop
    
    Call fctCreateEmail_MultipleRecipients(strTo, intcount, strCC, strSubject, strBody, blnSend, , True, strVotingOptions)

End Sub

