Public Function fct_Word_ModifyFooter(strFileName As String, strFooter As String) As Boolean
        
  'Dim wdApp As Word.Application
  Dim wdDoc As Word.Document

  On Error GoTo Err_fct_Word_ModifyFooter
  
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then 'Word not open
        Set wdApp = CreateObject("Word.Application")
    End If

    Set wdDoc = wdApp.Documents.Open(strFileName, , False)

    With wdDoc.Sections(1).Footers(wdHeaderFooterPrimary)
     If .Range.Text <> vbCr Then
     MsgBox .Range.Text
     Else
     MsgBox "Footer is empty"
     End If
    End With

lblSaveClose:
    wdDoc.Save
    wdDoc.Close
    
    'wdApp.DisplayAlerts = True
    
lblCleanup:

    Set xlWs = Nothing
  Set wdDoc = Nothing
  Set wdApp = Nothing

fct_Word_ModifyFooter = True
    
Exit Function
  

Err_fct_Word_ModifyFooter:
    
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 429 ' Word not open
            Resume Next
            ERR.Clear
        Case Else
            MsgBox "Error in fct_Word_ModifyFooter " & ERR.Number & ", Description: " & ERR.Description
            fct_WordDeleteWorksheetSTD = False
            Exit Function
        End Select
        GoTo lblCleanup
    End If



End Function
