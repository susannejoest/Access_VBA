Attribute VB_Name = "mdl_SqlServer"
Option Compare Database

Function ADOExecuteProcedure(strConnectionString As String, strExecProcedure As String)

   Dim ADOConn As New ADODB.Connection
   Dim ADOCom As New ADODB.Command

   On Error GoTo Err_Conn
    ADOConn.ConnectionString = strConnectionString '"ODBC;Description=Contiki PROD Write;DRIVER=SQL Server;SERVER=DESQL016.onetakeda.com\MSSQL010P;UID=Contiki;PWD=****;DATABASE=Contiki_app"
    
    ADOConn.Open
    ADOConn.CommandTimeout = 240 '120 seconds, run time ca 1 min 30 sec
    
    ADOConn.Execute strExecProcedure ' e.g. "EXEC Takeda_KeyWordSearch"
    'MsgBox "Keyword Search Procedure Done!"
    
    ADOConn.Close
    Set ADOConn = Nothing
    
    Exit Function
    
Err_Conn:

Select Case ERR.Number
    Case -2147217871 '[Microsoft][ODBC SQL Server Driver]Query timeout expired
        Stop
    Case Else
        Debug.Print ERR.Number
        Debug.Print ERR.Description
        Stop
End Select

End Function

Function fct_MergePersons(intPrsID As Integer, intPrsIDFinal As Integer)

    Dim cdb As DAO.Database, qdf As DAO.QueryDef
    Set cdb = CurrentDb
    Set qdf = cdb.CreateQueryDef("")
    ' get .Connect property from existing ODBC linked table
    qdf.Connect = cdb.TableDefs("TPersons").Connect
    qdf.Sql = "EXEC sp_MergePersons " & intPrsID & "," & intPrsIDFinal
    qdf.ReturnsRecords = False
    qdf.Execute dbFailOnError
    
    Set qdf = Nothing
    Set cdb = Nothing
    
End Function

Private Sub EntInactiveFlag_AfterUpdate()
    Dim strSQL As String
    
    'If me.EntInactiveFlag = False Then Exit Sub
    
    If MsgBox("You have set this entity to INACTIVE. Set all Signing Authority / Board etc. entries to Invalid?", vbYesNo) = vbYes Then
    
        strSQL = "UPDATE TList_SigningAuthority INNER JOIN TEntityListTakeda " _
        & "ON TList_SigningAuthority.SigEntityID = TEntityListTakeda.EntID " _
        & "SET TList_SigningAuthority.SigInactiveFlag = [EntInactiveFlag] " _
        & "WHERE TList_SigningAuthority.SigInactiveFlag=False " _
        & "AND SigEntityID = " & 1 'Me.EntID
    'AND EntInactiveFlag = True
        DoCmd.RunSQL (strSQL)
    Else
    
    End If
End Sub

Sub AddFieldNames()

    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim strColumnHeader As String
    Dim strColumnHeaderPrevious As String
    Dim strSQLUpdateFieldCodes As String
    
    'Blank out previous column headers
    DoCmd.RunSQL ("UPDATE tbl_EntityALL SET [ColumnHeader]='', [FieldCode]=''")
    
    'Update Country Fields
    
    strSQLUpdateFieldCodes = "UPDATE tbl_FieldCodes INNER JOIN tbl_EntityALL ON tbl_FieldCodes.FldFieldName= tbl_EntityALL.F1 " _
    & "SET tbl_EntityALL.FieldCode = [FldCode], tbl_EntityALL.ColumnHeader = [FldFieldName]"
    
    DoCmd.RunSQL (strSQLUpdateFieldCodes)
    
    strSQL = "SELECT * FROM tbl_EntityALL order by [AutoNum]"
    Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    
    strColumnHeader = ""
    
    Do Until rst.EOF

        'If rst!F1 Like "*Konz*" Then
        'Stop
        'End If
        
        If rst!ColumnHeader > "" Then
            strColumnHeader = rst!ColumnHeader
        Else
        
            
            If Not strColumnHeader = "#EMPTY" Then strColumnHeaderPrevious = strColumnHeader
        
            If IsNull(rst!F1) And IsNull(rst!F2) And IsNull(rst!F3) And IsNull(rst!F4) And IsNull(rst!F5) And IsNull(rst!F6) Then
                strColumnHeader = "#EMPTY"
            Else
                If strColumnHeader = "#EMPTY" Then strColumnHeader = strColumnHeaderPrevious
            End If

        End If
        
        rst.Edit
        rst!ColumnHeader = strColumnHeader
        'If rst!F1 = rst!ColumnHeader Then rst!ColumnHeader = rst!ColumnHeader & " Header"
        'If rst!F1 = "Name" Then rst!ColumnHeader = rst!ColumnHeader & " Header - Name"
        'If rst!F1 = "Company Name" Then rst!ColumnHeader = rst!ColumnHeader & " Header - Company Name"
        
        rst.Update
        
    rst.MoveNext
    Loop
    
    strSQLUpdateFieldCodes = "UPDATE tbl_FieldCodes INNER JOIN tbl_EntityALL ON tbl_FieldCodes.FldFieldName = tbl_EntityALL.ColumnHeader " _
    & "SET tbl_EntityALL.FieldCode = [FldCode]"
    
    DoCmd.RunSQL (strSQLUpdateFieldCodes)
    
    DoCmd.RunSQL ("UPDATE tbl_Countries INNER JOIN tbl_EntityALL ON tbl_Countries.CtyCountryName = tbl_EntityALL.F1 SET tbl_EntityALL.ColumnHeader = 'Country'")

End Sub

