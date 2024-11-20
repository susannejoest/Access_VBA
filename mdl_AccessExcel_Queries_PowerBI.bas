********************************
List Excel Power Queries into Access table
********************************

Sub ListExcelPowerQueries()
    Dim xlApp As Object ' Excel.Application
    Dim xlWb As Object ' Excel.Workbook
    Dim query As Object ' Excel.WorkbookQuery
    Dim conn As Object ' Excel.WorkbookConnection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim filePath As String
    Dim i As Integer
    
    ' Path to the Excel file it must not be open
    filePath = "C:\Users\p258097\OneDrive - Assicurazioni Generali S.p.A\Desktop\ProfitaS\Susanne\ProfitaS_Dashboard_wip 20 - Copy SJ.xlsm"
    
    ' Open Excel Application
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    If ERR.Number <> 0 Then
        MsgBox "Excel is not installed or could not be opened.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    xlApp.Visible = True
    Set xlWb = xlApp.Workbooks.Open(filePath)
    
    ' Open Access Table for data insertion
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("PowerQueries", dbOpenDynaset)
    
    ' Loop through Workbook Queries
    i = 0
    For Each query In xlWb.Queries
        rs.AddNew
        rs!queryName = query.Name
        rs!queryFormula = query.Formula
        rs.Update
        i = i + 1
        
    Next query
    
    ' Loop through Workbook Connections (if needed)
   ' For Each conn In xlWb.Connections
   '     rs.AddNew
   '     rs!QueryName = conn.Name
   '     rs!ConnectionString = conn.ODBCConnection.Connection
    '    rs.Update
   ' Next conn
    
    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    xlWb.Close SaveChanges:=False
    xlApp.Quit
    Set xlWb = Nothing
    Set xlApp = Nothing
    
    MsgBox i & " queries listed in the Access table 'PowerQueries'.", vbInformation
End Sub

********************************
Function WorksheetExists
********************************
 Error Handling:
            Case 1004 ' Tab already exists
            GoTo lblMoveNext ' Tab already exists
            xlWs.Delete

Function WorksheetExists(wb As Object, sheetName As String) As Boolean
    Dim ws As Object
    WorksheetExists = False
    
    ' Loop through all worksheets in the workbook
    For Each ws In wb.Sheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function
