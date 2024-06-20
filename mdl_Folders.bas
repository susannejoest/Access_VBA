Attribute VB_Name = "mdl_Folders"
Option Compare Database
 
Function cmdCreatefolders()

Dim strRootPath As String
Dim strQueryNameLstID As String 'LstID = column

strRootPath = "L:\Contiki\Projects\CMS\Issue_Folders"
strQueryNameLstID = "tbl_ListID"

Call fct_CreateFolders(strRootPath, strQueryNameLstID)

End Function

Function fct_CreateFolders(strRootPath As String, strQueryNameLstID As String)

Dim rs As DAO.Recordset

Set rs = CurrentDb.OpenRecordset(strQueryNameLstID)
        
     Do Until rs.EOF
     Call fctMkDirCreate(strRootPath & "\" & rs!LstID)
     
lblNext:
        rs.MoveNext
 Loop

    rs.Close
    Set rs = Nothing

End Function
