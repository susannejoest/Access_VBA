'Attribute VB_Name = "mdl_SharepointUpload"
'Option Compare Database

Public Function fct_UploadToSharepoint(strSourcePathAndFullFileName As String, strTargetFileNameNoPath As String, strSharepointUploadLibraryAddress As String) As Boolean

'file share A: is temporarily mapped as a workaround

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
