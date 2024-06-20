Attribute VB_Name = "mdl_FileList"
Option Compare Database


Function ListFilesInFolder(ByVal SourceFolderName As String, ByVal IncludeSubfolders As Boolean, Optional strFileType As String)
    
    'from file Listing database
    Dim db      As DAO.Database
    Dim sSQL    As String
     
    Set db = CurrentDb()
    
    On Error GoTo err_handler
    
    'Declaring variables
    Dim FSO As Object
    Dim SourceFolder As Object
    Dim SubFolder As Object
    Dim FileItem As Object
    Dim r As Long
       
    Dim File_Name As String
    Dim Parent_Folder As String
    Dim Path As String
    Dim File_Size As String
    Dim Date_Created As String
    Dim Date_LastModified As String
    Dim Date_LastAccessed As String
    
    'Creating object of FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    
    For Each FileItem In SourceFolder.Files
    
        Select Case FileItem.Type
            Case "Shortcut", "Data Base File", "TMP File", "Configuration settings", "Internet Shortcut", "Directory Query", "IconRemote Desktop Connection", "Icon"
                'do not list
            Case strFileType '"xlsx"
            'Date_LastAccessed
                'do events
       
                 'r = r + 1
         End Select
     
    Next FileItem
    
    'Getting files in sub folders
    If IncludeSubfolders Then
         For Each SubFolder In SourceFolder.SubFolders
            'Calling same procedure for sub folders
            ListFilesInFolder SubFolder.Path, True
         Next SubFolder
    End If
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
    
    Exit Function
    
err_handler:
    
    Debug.Print SourceFolderName
    Resume Next 'errors out if folder is empty or only has a system file
    

End Function

Function ImportDirListing(strPath As String, Optional strFilter As String)
    ' Author: CARDA Consultants Inc, 2007-01-19
    ' Copyright : The following may be altered and reused as you wish so long as the
    '             copyright notice is left unchanged (including Author, Website and
    '             Copyright).  It may not be sold/resold or reposted on other sites (links
    '              back to this site are allowed).
    '
    ' strPath = full path include trailing  ie:"c:windows"
    ' strFilter = extension of files ie:"pdf".  if you want to return
    '             a complete listing of all the files enter a value of
    '             "*" as the strFilter
    On Error GoTo Error_Handler
     
    Dim MyFile  As String
    Dim db      As DAO.Database
    Dim sSQL    As String
     
    Set db = CurrentDb()
     
    'Add the trailing  if it was omitted
    If Right(strPath, 1) <> "" Then strPath = strPath & ""
    'Modify the strFilter to include all files if omitted in the function
    'call
    If strFilter = "" Then strFilter = "*"
     
    'Loop through all the files in the directory by using Dir$ function
    MyFile = Dir$(strPath & "*." & strFilter)
    Do While MyFile <> ""
        'Debug.Print MyFile
        sSQL = "INSERT INTO DirectoryListing (File_Name) VALUES(""" & MyFile & """)"
        db.Execute sSQL, dbFailOnError
        'dbs.RecordsAffected 'could be used to validate that the
                                        'query actually worked
        MyFile = Dir$
    Loop
     
Error_Handler_Exit:
        On Error Resume Next
        Set db = Nothing
        Exit Function
     
Error_Handler:
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Number: " & ERR.Number & vbCrLf & _
               "Error Source: ImportDirListing" & vbCrLf & _
               "Error Description: " & ERR.Description, vbCritical, _
               "An Error has Occurred!"
        Resume Error_Handler_Exit
End Function




