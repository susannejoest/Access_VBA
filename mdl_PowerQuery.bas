Attribute VB_Name = "mdl_PowerQuery"
Option Compare Database

Function fct_ExcelApp_OpenWB_RefreshPowerQuery(strFilePathName As String, Optional strQueryName As String) ', blnRefreshIfCurrentMth As Boolean
'requires Microsoft Excel Reference


On Error GoTo ERR

Dim xlPowerQuery As Object
Dim xlWb As Excel.Workbook
Dim xlWs As Excel.Worksheet
Dim strMessage As String
Dim LastModDate_Query As Date

    Set xlApp = GetObject(, "Excel.Application") ' must trap error 429 Excel not open
    If xlApp Is Nothing Then 'Excel not open
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    xlApp.Visible = True
    
    ' If target workbook already open, use it
    For Each xlWb In xlApp.Workbooks
        If xlWb.FullName = strFilePathName Then
            ' If workbook is already open, set the workbook variable
            Set xlWb = xlWb
            GoTo lblNextStep
            Exit For
        End If
    Next xlWb
    
    ' If workbook is not open, open it
    Set xlWb = xlApp.Workbooks.Open(strFilePathName)

lblNextStep:

    ' Run the specific Power Query
    If strQueryName > "" Then
    'Stop 'assume ALL queries should be refreshed ,otherwise remove them
    
        For Each conn In xlWb.Connections
            Debug.Print conn.Name
            'If conn.Name <> "Abfrage - fact_PP_SS_MTD" Then GoTo lblNextConn
                ' Ensure it's a Power Query (by checking if it starts with "Query - " or "Abfrage - ")
                If conn.OLEDBConnection Is Nothing = False Then
                    If conn.Name = strQueryName Then
                        Debug.Print conn.Name
                        ' Get last modification date (if possible)
                         'LastModDate_Query = GetPowerQueryLastModifiedDate(pqQuery.Name, xlWb)
                         
                        ' Refresh only if last modified in the current month
                        'If Month(LastModDate) = Month(Date) And Year(LastModDate) = Year(Date) and blnRefreshIfCurrentMth Then
                            conn.Refresh
                            DoEvents
                            strMessage = strMessage & "Power Query " & conn.Name & " refreshed." & vbNewLine
                            Debug.Print strMessage
                            GoTo lblModelRefresh
                        'End If
                    End If
                End If
            Next conn
           

    Else
        
' Refresh ALL queries (no specific query name)

        For Each conn In xlWb.Connections
            Debug.Print conn.Name
            'If conn.Name <> "Abfrage - fact_PP_SS_MTD" Then GoTo lblNextConn
            ' Ensure it's a Power Query (by checking if it starts with "Query - " or "Abfrage - ")
            If conn.OLEDBConnection Is Nothing = False Then
                If Left(conn.Name, 8) = "Query - " Or Left(conn.Name, 10) = "Abfrage - " Then
                    
                    ' Get last modification date (if possible)
                    'LastModDate = GetLastRefreshDate(conn)
                    
                    ' Refresh only if last modified in the current month
                    If Month(LastModDate) = Month(Date) And Year(LastModDate) = Year(Date) Then
                        conn.Refresh
                        DoEvents
                        strMessage = strMessage & "Power Query " & conn.Name & " refreshed." & vbNewLine
                        Debug.Print strMessage
                    End If
                End If
            End If
lblNextConn:
        Next conn

        
    End If

lblModelRefresh:

    xlWb.Model.Refresh ' Refresh Power Pivot Model
    DoEvents
    
    xlApp.CalculateUntilAsyncQueriesDone 'wait until all refreshes completely done
    
    xlWb.Save
    xlWb.Close False 'prevents save prompt

lblExit:

Exit Function

ERR:
    Select Case ERR.Number
        Case 0 ' No error
        Case 91 'object variable or with block not set
            Stop
            Resume Next
        Case 429 ' Excel not open
            Resume Next
            ERR.Clear
        'Case 1004 ' Excel Worksheet already exists
        '    GoTo lblMoveNext ' Tab already exists
        '    xlWs.Delete
        Case Else
            Debug.Print ERR.Number & " " & ERR.Description
            Resume Next
    End Select
    
End Function



Sub RefreshPowerQueries()
    Dim xlWb As Workbook
    Dim conn As WorkbookConnection
    Dim objList As ListObject
    Dim xlWs As Worksheet
    Dim LastModDate As Date
    Dim strMessage As String
    
    ' Set workbook reference
    Set xlWb = ThisWorkbook
    
    ' Loop through all workbook connections
    For Each conn In xlWb.Connections
        ' Ensure it's a Power Query (by checking if it starts with "Query - " or "Abfrage - ")
        If conn.OLEDBConnection Is Nothing = False Then
            If Left(conn.Name, 8) = "Query - " Or Left(conn.Name, 10) = "Abfrage - " Then
                
                ' Get last modification date (if possible)
                LastModDate = GetLastRefreshDate(conn)
                
                ' Refresh only if last modified in the current month
                If Month(LastModDate) = Month(Date) And Year(LastModDate) = Year(Date) Then
                    conn.Refresh
                    DoEvents
                    strMessage = strMessage & "Power Query " & conn.Name & " refreshed." & vbNewLine
                    Debug.Print strMessage
                End If
            End If
        End If
    Next conn
    
    ' Notify user
    If strMessage <> "" Then
        MsgBox strMessage, vbInformation, "Power Queries Refreshed"
    Else
        MsgBox "No Power Queries needed refreshing.", vbInformation, "Status"
    End If
End Sub

' Function to get the last refresh date of a Power Query connection
Function GetLastRefreshDate(conn As WorkbookConnection) As Date
    On Error Resume Next ' Avoid errors if property is missing
    GetLastRefreshDate = conn.OLEDBConnection.RefreshDate
    On Error GoTo 0
End Function


' Function to get last refresh date of a Power Query
Function GetPowerQueryLastModifiedDate(QueryName As String, wb As Workbook) As Date

    Dim xml As Object
    Dim ns As Object
    Dim xmlDoc As Object
    Dim pqNode As Object
    Dim LastModDate As String
    
    ' Load Power Query metadata from Excel XML structure
    Set xml = wb.XmlMaps("PQ_QueriesMap").DataBinding
    Set xmlDoc = xml.xml
    Set ns = xmlDoc.DocumentElement.NamespaceURI

    ' Find the query node with the matching name
    For Each pqNode In xmlDoc.SelectNodes("//pq:Query", ns)
        If pqNode.Attributes.getNamedItem("name").Text = QueryName Then
            LastModDate = pqNode.Attributes.getNamedItem("lastRefresh").Text
            Exit For
        End If
    Next pqNode

    ' Convert the date string to VBA Date format
    If LastModDate <> "" Then
        GetPowerQueryLastModifiedDate = CDate(Left(LastModDate, 10))
    Else
        GetPowerQueryLastModifiedDate = 0 ' Default to 0 if no date found
    End If
End Function

Function SumTotalsInColumn(strWsName As String, strRangeAmt As String) As Integer

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalSum As Double
    
    ' Set the worksheet (change "Sheet1" to your actual sheet name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Sum all values in column B (change as needed)
    totalSum = 0 'Application.WorksheetFunction.Sum(ws.Range("B1:B" & lastRow))
    SumTotalsInColumn = totalSum
    ' Display the total sum
    'MsgBox "The sum of all totals in Column B is: " & totalSum, vbInformation, "Total Sum"
    
End Function




