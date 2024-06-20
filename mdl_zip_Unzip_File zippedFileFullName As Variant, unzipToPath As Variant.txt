Option Compare Database

Function CreateZipFile_ENTIREFOLDER(folderToZipPath As Variant, zippedFileFullName As Variant)
    ' usage: Call CreateZipFile("C:\Users\marks\Documents\ZipThisFolder\", "C:\Users\marks\Documents\NameOFZip.zip")
    
    Dim ShellApp As Object

    'Create an empty zip file
    Open zippedFileFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    'Copy the files & folders into the zip file
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(zippedFileFullName).CopyHere ShellApp.Namespace(folderToZipPath).Items
    
    'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
    On Error Resume Next
    Do Until ShellApp.Namespace(zippedFileFullName).Items.Count = ShellApp.Namespace(folderToZipPath).Items.Count
    Stop
        'Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0

End Function

Function UnzipAFile(zippedFileFullName As Variant, unzipToPath As Variant)
    
    Dim ShellApp As Object
    
    'Copy the files & folders from the zip into a folder
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(unzipToPath).CopyHere ShellApp.Namespace(zippedFileFullName).Items

End Function

Public Function zipFilesWithPassword()

    Dim source As String
    Dim target As String
    Dim password As String
    
    source = Chr$(34) & "C:\DEPTS\050\" & Chr$(34)
    target = Chr$(34) & "C:\DEPTS\050\050" & Chr$(34)
    
    password = Chr$(34) & "JORDAN" & Chr$(34)
    Shell ("C:\Program Files\WinZip\WINZIP32.EXE -min -a -sPASSWORD " & target & " " & source)

End Function


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

