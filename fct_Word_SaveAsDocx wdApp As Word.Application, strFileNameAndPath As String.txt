Public Function fct_Word_SaveAsDocx_VBAMODULES(wdApp As Word.Application, strFileNameAndPath As String)
', Optional blnWdAppVisible As Boolean, Optional blnDisplayAlerts As Boolean

    'wdApp.DisplayAlerts = blnDisplayAlerts
    'wdApp.Visible = blnWdAppVisible
     Dim wdDoc As Word.Document
     
     
    'Set wdDoc = wdApp.Documents.Open(strFileNameAndPath)
    Set wdDoc = wdApp.Documents.Open(strFileNameAndPath)
    strFileNameAndPath = strFileNameAndPath & "x" '.docx
    wdDoc.SaveAs FileName:=strFileNameAndPath, FileFormat:=wdFormatDocumentDefault '.docx

'Documents.Open FileName:="File1.doc", ConfirmConversions:= _
 '       False, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
  '      PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
  '      WritePasswordTemplate:="", Format:=wdOpenFormatAuto, XMLTransform:=""

Set wdDoc = Nothing

End Function
