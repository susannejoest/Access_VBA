Public Function FolderFromPath(strFullPath As String, intLevel As Integer) As String

    Dim I As Integer
    Dim intCountLevel As Integer
    
    intCountLevel = 0

    For I = Len(strFullPath) To 1 Step -1
        If Mid(strFullPath, I, 1) = "\" Then
            intCountLevel = intCountLevel = intCountLevel + 1
            If intCountLevel = intLevel Then
                FolderFromPath = Left(strFullPath, I)
            End If
            Exit For
        End If
    Next
    
End Function