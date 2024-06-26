Public Function fct_Word_App_GetOrCreate(blnWdAppVisible As Boolean, blnDisplayAlerts As Boolean) As Word.Application
               
  Dim wdApp As Word.Application
    On Error GoTo ERR
    DoCmd.SetWarnings False
 
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then 'Word not open
        Set wdApp = CreateObject("Word.Application")
    End If
    
    wdApp.Visible = blnWdAppVisible
    wdApp.DisplayAlerts = blnDisplayAlerts
    'Set wdAppPUBLIC = wdApp
    Set fct_Word_App_Public_VBAMODULES = wdApp
    Set wdAppPUBLIC = wdApp
    
lblSaveClose:
   
lblCleanup:
    
Exit Function
  
ERR:
    
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 429 ' Word not open
            Resume Next
            ERR.Clear
        Case Else
            'fct_Word_App = False
                Stop
                Resume Next 'errors out if folder is empty or only has a system file
    
            Exit Function
        End Select
        GoTo lblCleanup
    End If

End Function
