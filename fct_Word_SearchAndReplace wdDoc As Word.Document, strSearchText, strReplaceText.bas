Function fct_Word_SearchAndReplace(wdDoc As Word.Document, strSearchText, strReplaceText)
    
    If strSearchText Is Null Then strSearchText = "{the Company}"
    If strReplaceText Is Null Then strReplaceText = strName
        With wdDoc.Content.Find
            .Text = strSearchText
            .Replacement.Text = strReplaceText
            .Execute Replace:=wdReplaceOne
        End With
End Function
