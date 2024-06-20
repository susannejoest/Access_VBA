Attribute VB_Name = "mdl_StringFunctions"
Option Compare Database

Public Function fctString_MakeAlphaNumeric(inputStr As String, blnNumbers As Boolean, blnSpace As Boolean)
'removes all but letters AND numbers

    Dim ascVal As Integer, originalStr As String, newStr As String, counter As Integer, trimStr As String
    
    On Error GoTo Err_Stuff
    ' send to error message handler
    If inputStr = "" Then Exit Function
    ' if nothing there quit
    trimStr = Trim(inputStr)
    ' trim out spaces
    newStr = ""
    ' initiate string to return
    For counter = 1 To Len(trimStr)
        ' iterate over length of string
        ascVal = Asc(Mid$(trimStr, counter, 1))
        'Debug.Print ascVal
        'Debug.Print Mid$(trimStr, counter, 1)
        'Debug.Print "**************"
        ' find ascii vale of string
        Select Case ascVal
            Case 32  ' 32 = space  '48 To 57, 65 To 90 ,
                If blnSpace Then newStr = newStr & Chr(ascVal)
            Case 48 To 57 ' number, 48 = 0
                If blnNumbers Then newStr = newStr & Chr(ascVal)
            'Case 63 '? e.g. Chinese Char
            Case 65 To 90, 97 To 122 ' 65=lower case, 97 = upper case letter
                ' if value in case then acceptable to keep
                newStr = newStr & Chr(ascVal)
                ' add new value to existing new string
            Case Else '58 - 64
                'Debug.Print ascVal
        End Select
    Next counter
    ' move to next character
    fctString_MakeAlphaNumeric = newStr
    ' return new completed string
    
Exit Function
Err_Stuff:
    ' handler for errors
    MsgBox ERR.Number & " " & ERR.Description
End Function

Function fctValidFileName(strFileNameOld As String, Optional strReplaceChar) As String
'"\/:*?<>|[]"""
'single quotes are NOT removed, use file name 2

If IsNull(strReplaceChar) Or IsMissing(strReplaceChar) Then strReplaceChar = "_"

    fctValidFileName = Replace(strFileNameOld, "/", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    fctValidFileName = Replace(strFileNameOld, "\", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    fctValidFileName = Replace(strFileNameOld, "&", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, "<", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, ">", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, ";", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, """", "'", 1, -1, vbTextCompare) ' single quote was previously underscore "
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, ":", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, "*", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, "|", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, "[", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
    fctValidFileName = Replace(strFileNameOld, "]", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName
    
End Function

Function fctValidFileName2(strFileNameOld As String, Optional strReplaceChar) As String
'remove single quote

If IsNull(strReplaceChar) Then strReplaceChar = "_"

    fctValidFileName2 = fctValidFileName(strFileNameOld, strReplaceChar)

    
    fctValidFileName2 = Replace(strFileNameOld, ".", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName2
    
    fctValidFileName2 = Replace(strFileNameOld, "'", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName2
    
    fctValidFileName2 = Replace(strFileNameOld, "__", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName2
    
End Function

Function fctValidFileName_ColonSeparatedString(strFileNameOld As String, Optional strReplaceChar) As String
'"\/:*?<>|[]"""
'single quotes are NOT removed, use file name 2
'& not removed

If IsNull(strReplaceChar) Then strReplaceChar = "_"
    'fctValidFileName = Replace(strFileNameOld, ";", strReplaceChar, 1, -1, vbTextCompare)
    'strFileNameOld = fctValidFileName
        'fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "&", strReplaceChar, 1, -1, vbTextCompare)
    'strFileNameOld = fctValidFileName_ColonSeparatedString
    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "/", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "\", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString

    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "<", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, ">", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    

    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, """", strReplaceChar, 1, -1, vbTextCompare) ' single quote was previously underscore "
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, ":", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
     fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, ".", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "*", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "|", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "[", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
    fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "]", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
        fctValidFileName_ColonSeparatedString = Replace(strFileNameOld, "'", strReplaceChar, 1, -1, vbTextCompare)
    strFileNameOld = fctValidFileName_ColonSeparatedString
    
End Function

Function fctValidFileName_ReplaceHyphen(strFileNameOld As String, Optional strReplaceChar) As String

If IsNull(strReplaceChar) Then strReplaceChar = "_"

        fctValidFileName_ReplaceHyphen = Replace(strFileNameOld, "'", strReplaceChar, 1, -1, vbTextCompare)

End Function
'"\/:*?<>|[]"""
'single quotes are NOT removed, use file name 2
'& not removed



Function fctStripFillChar(strFileNameOld As String) As String

    strFileNameOld = Replace(strFileNameOld, "/", "", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, "&", "", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, " ", "", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, ".", "", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, "-", "", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, "(", "", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, ")", "", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, "ue", "ü", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, "ae", "ä", 1, -1, vbTextCompare)
    strFileNameOld = Replace(strFileNameOld, "oe", "ö", 1, -1, vbTextCompare)
    
    strFileNameOld = Replace(strFileNameOld, ":", "-", 1, -1, vbTextCompare)
     
    strFileNameOld = StrConv(strFileNameOld, 1) 'all upper case
    
    fctStripFillChar = strFileNameOld
    
End Function

Function fctReplaceAmp(strFileNameOld As String) As String

    strFileNameOld = Replace(strFileNameOld, "&amp;", "&", 1, -1, vbTextCompare)
    fctReplaceAmp = strFileNameOld

End Function


Public Function SplitString_20Col(strParent_Folder As String, delimiter As String) As Variant
    
    Dim strArr() As String
    
    Dim L01 As String
    Dim L02 As String
    Dim L03 As String
    Dim L04 As String
    Dim L05 As String
    Dim L06 As String
    Dim L07 As String
    Dim L08 As String
    Dim L09 As String
    Dim L10 As String
    Dim L11 As String
    Dim L12 As String
    Dim L13 As String
    Dim L14 As String
    Dim L15 As String
    Dim L16 As String
    Dim L17 As String
    Dim L18 As String
    Dim L19 As String
    Dim L20 As String

    
    Dim i As Integer
    
    i = 0
    
    strArr = Split(strParent_Folder, delimiter)
    
    For i = 0 To UBound(strArr)
    
        If i = 1 Then
            L00 = strArr(0)
            GoTo lblNext
        End If
        
        If i = 1 Then
            L00 = strArr(1)
            GoTo lblNext
        End If
        
        If i = 2 Then
            L00 = strArr(2)
            GoTo lblNext
        End If
        
        If i = 3 Then L03 = strArr(3)
        If i = 4 Then L04 = strArr(4)
        If i = 5 Then L05 = strArr(5)
        If i = 6 Then L06 = strArr(6)
        If i = 7 Then L07 = strArr(7)
        If i = 8 Then L08 = strArr(8)
        If i = 9 Then L09 = strArr(9)
        If i = 10 Then L10 = strArr(10)
        If i = 11 Then L11 = strArr(11)
        If i = 12 Then L12 = strArr(12)
        If i = 13 Then L13 = strArr(13)
        If i = 14 Then L14 = strArr(14)
        If i = 15 Then L15 = strArr(15)
        If i = 16 Then L16 = strArr(16)
        If i = 17 Then L17 = strArr(17)
        If i = 18 Then L18 = strArr(18)
        If i = 19 Then L19 = strArr(19)
        If i = 20 Then L20 = strArr(20)
        
lblNext:
    Next i
    

   ' count = count - 1 'zero-based
    'If UBound(strArr) >= count Then

        
       ' SplitString = strArr(count)

    
End Function

Function WordsInString_Count(ByVal S, strSeparator As String) As Integer
      ' Counts the words in a string that are separated by commas.
    'needs to be in a normal module, not class module to call via immediate pane
    
      Dim WC As Integer, Pos As Integer
         If VarType(S) <> 8 Or Len(S) = 0 Then
           WordsInString_Count = 0
           Exit Function
         End If
         
         WC = 1
         Pos = InStr(S, strSeparator) '"#"
         
         Do While Pos > 0
           WC = WC + 1
           Pos = InStr(Pos + 1, S, strSeparator)
         Loop
         
         WordsInString_Count = WC
         
End Function
      
Function WordsInString_GetWordByIndex(ByVal S, strSeparator As String, Indx As Integer)
      ' Returns the nth word in a specific field.

      Dim WC As Integer, Count As Integer, SPos As Integer, EPos As Integer
         WC = WordsInString_Count(S, strSeparator)
         
         If Indx < 1 Or Indx > WC Then
           WordsInString_GetWordByIndex = Null
           Exit Function
         End If
         
         Count = 1
         SPos = 1
         
         For Count = 2 To Indx
           SPos = InStr(SPos, S, strSeparator) + 1
         Next Count
         
         EPos = InStr(SPos, S, strSeparator) - 1
         If EPos <= 0 Then EPos = Len(S)
         
         WordsInString_GetWordByIndex = Trim(Mid(S, SPos, EPos - SPos + 1))
         
End Function
