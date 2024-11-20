*********************************
GetFileLastModifiedDate (show in Excel cell when other workbook last updated)
*********************************

Function GetFileLastModifiedDate(filePath As String) As String
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If fileSystem.FileExists(filePath) Then
        GetFileLastModifiedDate = fileSystem.GetFile(filePath).DateLastModified
    Else
        GetFileLastModifiedDate = "File Not Found"
    End If
    On Error GoTo 0
    
    Set fileSystem = Nothing
End Function
