Sub fct_Word_ChangeDocumentAuthorName_To_UserName(wdApp, wdDoc, strUserName As String)
On Error GoTo ERR
'Dim strUserName As String
  'strUserName = Application.UserName
  'Application.UserName = InputBox("Enter new author")
  With wdDoc 'Word.Document
    .BuiltinDocumentProperties("Author") = strUserName 'wdapp.UserName
    .Save
  End With
  
  wdApp.UserName = strUserName 'wdApp = Word.Application
  
  wdDoc.Close
lbl_Exit:
  Exit Sub
  
ERR:
Debug.Print ERR.Number & ": " & ERR.Description
Stop

End Sub
