Public Declare PtrSafe Function Sleep Lib "kernel32" ( _
ByVal dwMilliseconds As Long)  ' for zipping function

Function fct_ZipSingleFile(ZipFile As String, InputFile As String)
' must include mdl_apiFunctions Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

On Error GoTo ErrHandler
Dim FSO As Object 'Scripting.FileSystemObject
Dim oApp As Object 'Shell32.Shell
Dim oFld As Object 'Shell32.Folder
Dim oShl As Object 'WScript.Shell
Dim i As Long
Dim l As Long

Set FSO = CreateObject("Scripting.FileSystemObject")
If Not FSO.FileExists(ZipFile) Then
'Create empty ZIP file
FSO.CreateTextFile(ZipFile, True).Write _
"PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
End If

Set oApp = CreateObject("Shell.Application")
Set oFld = oApp.Namespace(CVar(ZipFile))
i = oFld.Items.Count
oFld.CopyHere (InputFile)

Set oShl = CreateObject("WScript.Shell")

'Search for a Compressing dialog
Do While oShl.AppActivate("Compressing...") = False
If oFld.Items.Count > i Then
'There's a file in the zip file now, but
'compressing may not be done just yet
Exit Do
End If
If l > 30 Then
'3 seconds has elapsed and no Compressing dialog
'The zip may have completed too quickly so exiting
Exit Do
End If

DoEvents
Sleep 100
l = l + 1
Loop

' Wait for compression to complete before exiting
Do While oShl.AppActivate("Compressing...") = True
DoEvents
Sleep 100
Loop

ExitProc:
On Error Resume Next
Set FSO = Nothing
Set oFld = Nothing
Set oApp = Nothing
Set oShl = Nothing
Exit Function
ErrHandler:
Select Case ERR.Number
Case Else
MsgBox "Error " & ERR.Number & _
": " & ERR.Description, _
vbCritical, "Unexpected error"
End Select
Resume ExitProc
Resume
End Function
