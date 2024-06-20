Attribute VB_Name = "mdl_Recordset"
Option Compare Database

Function fct_DAORecordSet_TblQry(strTblQryName As String, Optional strWHERE As String)

On Error GoTo err_handler

Dim rs As DAO.Recordset
Dim strSQL As String

strSQL = "SELECT * FROM " & strTblQryName & IIf(strWHERE > "", " WHERE " & strWHERE, "")
     Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
         
     Do Until rs.EOF
         Debug.Print rs(1).Value
         'rs!ArtikelID
         'rs![Field With Space]
         
lblMoveNext:
         rs.MoveNext
     Loop
     
lblCleanup:
    rs.Close
    Set rs = Nothing
    'Set db = Nothing
    
Exit Function

err_handler:
    If ERR.Number <> 0 Then
        Select Case ERR.Number
            Case 0
            Case Else
            MsgBox "Unhandled Error in fct_Recordset: " & ERR.Number & " " & ERR.Description
                Stop
        End Select
    End If

GoTo lblCleanup

End Function











