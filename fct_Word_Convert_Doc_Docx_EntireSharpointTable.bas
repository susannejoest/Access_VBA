Public Function fct_Convert_Word_Doc_Docx_EntireSharepointTable()
'Word module
On Error GoTo ERR
    ' This pulls the data from a SharePoint list. Don't forget a reference
    ' to the MS ActiveX Objects in "References".
     
    ' Variable declarations.
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConn As String
    Dim sSQL As String
    Dim wdDoc As Word.Document
    Dim blnWdAppVisible As Boolean
    
    blnWdAppVisible = True
    ' Build the connection string: use the ACE engine, the database is the URL if your site
    ' and the GUID of your SharePoint list is the LIST.
    sConn = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;" & _
    "DATABASE=" & sSHAREPOINT_SITE & ";" & _
    "LIST=" & sDEMAND_ROLE_GUID & ";"
     
    ' Create some new objects.
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
     
    ' Open the connection.
    With cn
    .ConnectionString = sConn
    .Open
    End With
     
    ' Build your SQL the way you would any other. You can add a WHERE clause
    ' if you need to filter on some criteria.
    'sSQL = "SELECT [HasElectronicSignatureClause] FROM [Templates] WHERE ID = " & intSPListID
    
    Call fct_Word_App_Public_VBAMODULES(blnWdAppVisible, True)
    
    sSQL = "SELECT * FROM tbl_SharepointDocs WHERE [File Type]=""doc"" " 'tbl_SharepointDocs is a linked Sharepoint Table

    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
    Do While Not rs.EOF
         Set wdDoc = wdApp.Documents.Open(rs![Encoded Absolute URL], ReadOnly:=True)
         wdDoc.SaveAs2 FileName:=rs![Encoded Absolute URL] & "x", FileFormat:=wdFormatXMLDocument
         'strPath & strFileName
           ' wddoc.SaveAs2 (strpath & strfilename, ".docx") '   .wdActiveDoc.SaveAs "testt.docx"
        rs.MoveNext
    Loop
'open recordset from


    ' Open up the recordset.
    
    'Debug.Print rs![HasElectronicSignatureClause]
    
lblCleanup:
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing

ERR: Stop
Debug.Print ERR.Description

End Function
