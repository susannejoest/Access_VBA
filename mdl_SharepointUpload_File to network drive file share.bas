/* README */

' This function Uploads a specific file from a local file share to Sharepoint
' The function only works if 
'1. you are able to add the sharepoint site you want to upload to to your Trusted Sites List 
' If your administrator has disabled this then the procedure won't work. Synched sharpeoint folders are also not an option, they do not work here.
' OR
'2. you still have the old Internet Explorer and can just map sharepoint sites out of the box
' alternatively you can use the approach in my other file copy procedure, which bypasses these issues: CopyFileLocalShareToSharepointWithTimerLoop

/* FUNCTION */

'Attribute VB_Name = "mdl_SharepointUpload"
'Option Compare Database

Public Function fct_UploadToSharepoint(strSourcePathAndFullFileName As String, strTargetFileNameNoPath As String, strSharepointUploadLibraryAddress As String) As Boolean
' strSourcePathAndFullFileName = full file path and name, e.g. "R:\FIN_Reporting_Analytics\Tagetik\Tagetik_Report_Download\Clearance\Report1.xlsx"
' strTargetFileNameNoPath = Target file name in Sharepoint, e.g. Report1.xlsx
' strSharepointUploadLibraryAddress = sharepoint site to upload to, e.g. \\my.sharepoint.com@SSL\DavWWWRoot\sites\mysite\Shared Documents\"

'file share A: is temporarily mapped as a workaround (only works if you are able to map sharepoint sites as file shares, if it is not disabled)

On Error GoTo Err_fct_UploadToSharepoint
    Dim objNet As Object 'CreateObject("WScript.Network")
    Dim FS As Object ' CreateObject("Scripting.FileSystemObject")
 
    Set objNet = CreateObject("WScript.Network")
    Set FS = CreateObject("Scripting.FileSystemObject")
    objNet.MapNetworkDrive "A:", strSharepointUploadLibraryAddress
    
    CreateObject("Shell.Application").Open (strSharepointUploadLibraryAddress) 'SP page must be open because of authentication or upload will fail
     'objShell.Shell.Open (ssfPERSONAL)
             
lblUpload:
    If FS.FileExists(strSourcePathAndFullFileName) Then
        FS.CopyFile strSourcePathAndFullFileName, "A:" & strTargetFileNameNoPath, True 'True = overwrite yes, otherwise it will not overwrite if false
    End If
     
    fct_UploadToSharepoint = True ' function succeeded

lblCleanup:
    objNet.RemoveNetworkDrive "A:" 'if this is not removed then there will be an error next time this function runs so the cleanup must run in any case
    Set objNet = Nothing
    Set FS = Nothing
    
Exit Function

Err_fct_UploadToSharepoint:
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case -2147024811 'The local device name is already in use, A already mapped
            objNet.RemoveNetworkDrive "A:"
            Resume Next
        Case Else
            MsgBox "Error: " & ERR.Number & ", " & ERR.Description
            fct_UploadToSharepoint = False
            Stop
            GoTo lblCleanup
        End Select
    End If

End Function



