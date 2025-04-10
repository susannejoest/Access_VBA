Option Compare Database

Private Declare PtrSafe Function CopyFileA Lib "kernel32" (ByVal ExistingFileName As String, _
  ByVal NewFileName As String, ByVal FailIfExists As Long) As Long


Function fct_CopyFile_IfExists(strFilePathSRC As String, strFileNameSRC As String, strFilePathDST As String, strFileNameDST As String, blnOverwrite As Boolean, blnCopyOnlyIfNotExists As Boolean)
' 10-Apr-2025

    Dim FSO As Object
    
    If FSO.FileExists(strFilePathDST & strFileNameDST) Then

       If blnCopyOnlyIfNotExists Then 'blnCopyOnlyIfNotExists = True
       
            'do nothing
            
        Else: blnCopyOnlyIfNotExists = False
        
            FSO.CopyFile strFilePathSRC & strFileNameSRC, strFilePathDST & strFileNameDST
            
        End If
        
    End If
    
    Set FSO = Nothing
    
End Function


Function fctChkMkDir(strRootPath As String, Optional datDate As Date, Optional strMode As String) As String

On Error GoTo Error_fctChkMkDir

Dim fsoFolder As Variant
Dim strPath As String

If Not Right(strRootPath, 1) = "\" Then strRootPath = strRootPath & "\"
'If Not Dir(strRootPath) Then MkDir (strRootPath)

If Not datDate = "12:00:00 AM" Then '12:00:00 AM is the default if no value is supplied
    strPath = strRootPath & fctDateFolder(datDate)
    strPath = strPath & "\"
Else
    strPath = strRootPath
End If


    
    fsoFolder = Dir(strPath, vbDirectory)
    If fsoFolder = "" Then
    fsoFolder = fctMkDirPlus(strPath)     ' Make new directory or folder.
    'fctMkDirPlus (strPath)
    End If
    
    fsoFolder = Dir(strPath, vbDirectory)
    If Not fsoFolder = "" Then fctChkMkDir = strPath

Exit Function

Error_fctChkMkDir:
    Select Case ERR.Number
        Case 76 'folder does not exist
        MsgBox "folder does not exist!!!"
    End Select


End Function

Function fctDateFolder(datDate As Date, Optional strMode As String) As String
If Not IsDate(datDate) Then fctDateFolder = ""

Dim strY As String
Dim strMY As String
Dim strD As String

    strY = Format(datDate, "yyyy")
    strMY = Format(datDate, "mm_yyyy")
    strD = Format(datDate, "mm_dd_yy")
    
    'If Not Right(strRootPath, 1) = "\" Then strRootPath = strRootPath & "\"
    'If Not Dir(strRootPath) Then MkDir (strRootPath)
    
    Select Case strMode
        Case modYMYD
            fctDateFolder = strY & "\" & strMY & "\" & strD
        Case modMYD
            fctDateFolder = strMY & "\" & strD
        Case modD
            fctDateFolder = strD
        Case Else   'default is this:
            fctDateFolder = strY & "\" & strMY & "\" & strD
    End Select

End Function

Function fctMkDirCreate(strPath As String) As String

Dim fsoFolder As Variant

    fsoFolder = Dir(strPath, vbDirectory)
    If Not fsoFolder = "" Then  'directory already exists!!!
        fctMkDirCreate = strPath
        Exit Function
    End If
    
    'create directory if it does not exist
 
    MkDir (strPath) 'create directory
    
    fsoFolder = Dir(strPath, vbDirectory)
    If Not fsoFolder = "" Then fctMkDirCreate = strPath

End Function

Function fctMkDirPlus(strPath As String) As String
Dim arrPaths(100) As String
Dim intcount As Integer
Dim fsoFolder As Variant
Dim strPathNew As String
strPathNew = strPath
intcount = 1

    fsoFolder = Dir(strPath, vbDirectory)
    If Not fsoFolder = "" Then  'no problem, directory already exists!!!
        fctMkDirPlus = fsoFolder
        Exit Function
    End If
    
    'arrPaths(intCount) = strPath
    
    Do While fsoFolder = ""
        arrPaths(intcount) = strPathNew
        intcount = intcount + 1
        strPathNew = fctGetPriorFolder(strPathNew)
        fsoFolder = Dir(strPathNew, vbDirectory)

    Loop
    intcount = intcount - 1
    
    Do Until intcount = 0
        MkDir (arrPaths(intcount))
        intcount = intcount - 1
    Loop
    
    fsoFolder = Dir(strPath, vbDirectory)
    If Not fsoFolder = "" Then fctMkDirPlus = strPath

End Function


Function fctGetPriorFolder(strPath As String) As String
    Dim strPathPrior As String

    'strip away "\"
    Do Until Not Right(strPath, 1) = "\"
        If Not Len(strPath) > 1 Then Exit Do
        strPath = Mid(strPath, 1, Len(strPath) - 1)
    Loop

    Do Until Right(strPath, 1) = "\"
        If Not Len(strPath) > 1 Then Exit Do
        strPath = Mid(strPath, 1, Len(strPath) - 1)
    Loop

    If Len(strPath) > 1 Then
        fctGetPriorFolder = strPath
        Else
        fctGetPriorFolder = strPath
    End If

End Function


  
Public Function Copy(FileSrc As String, FileDst As String, Optional NoOverWrite As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Copy = CopyFileA(FileSrc, FileDst, NoOverWrite) = 1

ExitRoutine:
    On Error Resume Next
    Exit Function
ErrorHandler:
    With ERR
        Select Case .Number
            Case Else
                MsgBox .Number & vbCrLf & .Description, vbInformation, "Error - Copy"
        End Select
    End With
    'Resume 0
    Resume ExitRoutine
End Function

Public Function fct_CopyFile(strSourceFileName As String, strTargetFileName As String) As Boolean
     ' better use copy function, this one has problems with permissions
     On Error GoTo Err_CopyFile
     
     FileCopy strSourceFileName, strTargetFileName
    fct_CopyFile = True
    
Exit Function
Err_CopyFile:

Select Case ERR.Number
    Case 0
    Case 33
    Case 70 'Permission denied
        If MsgBox("File is open, close it first before it can be copied! Retry?", vbYesNo) = vbYes Then
            
            Resume
        Else
            fct_CopyFile = False
            MsgBox "File could not be copied!"
            Exit Function
        End If
        
    Case Else
    Stop
End Select

End Function

Function fct_DeleteKillFile(strPathFileName) As Boolean

    Dim fsoFolder As Variant
' check if final xls file already exists, if so, delete
  fsoFolder = Dir(strPathFileName, vbDirectory)
    If fsoFolder <> "" Then
        Kill strPathFileName
    End If
    
End Function

Function RenameFile(folderName As String, searchFileName As String, renameFileTo As String)
'Call renamefile("V:\Templates\PLG\LOC_CH","-bla","")
 Dim FSO, folder, file

 'folderName = "C:\TEMP\LOC_CH"
 'todaysDate = Date

 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set folder = FSO.GetFolder(folderName)
 
 For Each file In folder.Files
     If (file Like "*-bla*") Then
          file.Name = Replace(file.Name, searchFileName, renameFileTo)
     End If
 Next

End Function

