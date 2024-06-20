Attribute VB_Name = "mdl_AccessTableFunctions"
Option Compare Database

Function fct_DeleteEntityTbl()

    Dim t As TableDef
    For Each t In CurrentDb.TableDefs
        If t.Name Like "tbl_Entity_*" Or t.Name Like "*_ImportError*" Then
            DoCmd.DeleteObject acTable, t.Name
        End If
    Next
    
End Function
