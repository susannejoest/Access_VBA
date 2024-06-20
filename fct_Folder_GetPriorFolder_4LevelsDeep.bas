Option Compare Database

Function test()
Call fctGetPriorFolder_WithSeparator("/CSEE/Belgium/Younity", 3)
End Function



Function fctGetPriorFolder_WithSeparator(strPath As String, intLevel As Integer, Optional strSeparator As String) As String

On Error GoTo ERR

    Dim strPathPrior As String
    Dim intCount As Integer
    Dim isSeparator As Boolean
    Dim isMaxIntCount As Boolean
    Dim intSubtract As Integer
    Dim strFolderName_Output As String
    
    Dim intLevelCount As Integer
    
    Dim strLevel_1_Folder As String
    Dim intLevel_1_Separator As Integer
 
    Dim strLevel_2_Folder As String
    Dim intLevel_2_Separator As Integer
  
    Dim strLevel_3_Folder As String
    Dim intLevel_3_Separator As Integer
    
    Dim strLevel_4_Folder As String
    Dim intLevel_4_Separator As Integer
    
    intCount = 1
    intLevelCount = 0
    
    If strSeparator = "" Then strSeparator = "/"
    If Left(strPath, 1) = strSeparator Then strPath = Mid(strPath, 2, Len(strPath)) 'leading separator, strip off
    'strip away "\"
    Do Until Len(strPath) < intCount '21
    'Not Right(strPath, intcount) = strSeparator '"\"
        If Right(Left(strPath, intCount), 1) = strSeparator Then isSeparator = True Else isSeparator = False
        If Len(strPath) = intCount Then isMaxIntCount = True Else isMaxIntCount = False
        If isMaxIntCount Then intSubtract = 0 Else intSubtract = 1
        
        
        If isSeparator = True Or isMaxIntCount Then 'either separator at end, or not
            intLevelCount = intLevelCount + 1
            
            Select Case intLevelCount
            Case 1
                intLevel_1_Separator = intCount
                strLevel_1_Folder = Left(strPath, intCount - intSubtract) 'Mid(strPath, intCount, Len(strPath) - 1)
                If intLevel = intLevelCount Then strFolderName_Output = strLevel_1_Folder
            Case 2
                intLevel_2_Separator = intCount
                strLevel_2_Folder = Mid(strPath, intLevel_1_Separator + 1, intLevel_2_Separator - intLevel_1_Separator - intSubtract)
                If intLevel = intLevelCount Then strFolderName_Output = strLevel_2_Folder
            Case 3
                intLevel_3_Separator = intCount
                strLevel_3_Folder = Mid(strPath, intLevel_2_Separator + 1, intLevel_3_Separator - intLevel_2_Separator - intSubtract)
                If intLevel = intLevelCount Then strFolderName_Output = strLevel_3_Folder
            Case 4
                intLevel_4_Separator = intCount
                strLevel_4_Folder = Mid(strPath, intLevel_3_Separator + 1, intLevel_4_Separator - intLevel_3_Separator - intSubtract)
                If intLevel = intLevelCount Then strFolderName_Output = strLevel_4_Folder
            End Select
            
        
        Else: isSeparator = False
        End If
        
        'If isSeparator Then
            'If Not Len(strPath) > 1 Then Exit Do
            'strPath = Mid(strPath, 1, Len(strPath) - 1)
            'intLevelCount = intLevelCount + 1
        'End If
        
        intCount = intCount + 1
    Loop

    'Do Until Right(strPath, 1) = strSeparator
    '    If Not Len(strPath) > 1 Then Exit Do
    '    strPath = Mid(strPath, 1, Len(strPath) - 1)
    'Loop

    'If Len(strPath) > 1 Then
    '    fctGetPriorFolder_WithSeparator = strPath
    '    Else
    '    fctGetPriorFolder_WithSeparator = strPath
    'End If
    
    fctGetPriorFolder_WithSeparator = strFolderName_Output ' returns null if that level does not exist
    
Exit Function
ERR: Stop

End Function

