Function StringCountOccurrences(strText As String, strFind As String, _
                                Optional lngCompare As VbCompareMethod) As Long

    Dim lngPos As Long
    Dim lngTemp As Long
    Dim lngCount As Long
    
        If Len(strText) = 0 Then Exit Function
        If Len(strFind) = 0 Then Exit Function
        lngPos = 1
        
        Do
            lngPos = InStr(lngPos, strText, strFind, lngCompare)
            lngTemp = lngPos
            If lngPos > 0 Then
                lngCount = lngCount + 1
                lngPos = lngPos + Len(strFind)
            End If
        Loop Until lngPos = 0
        StringCountOccurrences = lngCount
        
End Function