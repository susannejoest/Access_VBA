****************************************
EXCEL APP - Create or get Excel App
****************************************
Function fct_ExcelApp(strFilePath As String, xlWb As String)

    Set xlApp = GetObject(, "Excel.Application") ' must trap error 429 Excel not open
    If xlApp Is Nothing Then 'Excel not open
        Set xlApp = CreateObject("Excel.Application")
    End If

    ' If target workbook already open, use it
    For Each xlWb In xlApp.Workbooks
        If xlWb.FullName = filePath Then
            ' If workbook is open, set the workbook variable
            Set xlWb = xlWb
            GoTo lblNextStep
            Exit For
        End If
    Next xlWb
    
    ' If workbook is not open, open it
    Set xlWb = xlApp.Workbooks.Open(filePath)

lblNextStep:

    Select Case ERR.Number
        Case 0 ' No error
        Case 91 'object variable or with block not set
            Stop
            Resume Next
        Case 429 ' Excel not open
            Resume Next
            ERR.Clear
        Case 1004 ' Excel Worksheet already exists
            GoTo lblMoveNext ' Tab already exists
            xlWs.Delete
        Case Else
            Debug.Print ERR.Number & " " & ERR.Description
            Resume Next
    End Select

End Function

************
EXPORT EXCEL to Access?
*****************

        Attribute VB_Name = "mdl_CreateExcel"
        Option Compare Database
        Option Explicit
        
        Public Const blnxlAppVisible As Boolean = False
        Public xlAppWasOpen As Boolean
        
        Public Function fct_ExportExcel(strQueryName As String, strFileName As String) As Boolean
        
                DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strQueryName, strFileName, True
        
        End Function
        
        Public Function fct_CreateQueryDefSTD(strQueryName As String, strSQL As String) As Boolean
        
        On Error GoTo Err_fct_CreateQueryDef
        
                CurrentDb.QueryDefs.Delete strQueryName
                CurrentDb.CreateQueryDef strQueryName, strSQL
             
        fct_CreateQueryDefSTD = True
        
        Exit Function
        
        Err_fct_CreateQueryDef:
            If ERR.Number <> 0 Then
                Select Case ERR.Number
                    Case 3265 ' qry_temp cannot be deleted since it does not exist, Item not found in this collection
                        ERR.Clear
                        Resume Next
                    Case 0 'spreadsheet is open
                    Case Else
                    MsgBox "Unhandled Error in fct_CreateQueryDefSTD No: " & ERR.Number & " " & ERR.Description
                    Exit Function
        
                End Select
            End If
                
        End Function

*****************
CREATE EMPTY WORKBOOK
*****************

Public Function fct_CreateEmptyExcelWorkbookSTD(strFileName As String) As Boolean
               
  Dim xlApp As Excel.Application
  Dim xlWb As Excel.Workbook
  
  On Error GoTo Err_fct_ExcelCreateWorkbook
  
    Set xlApp = GetObject(, "Excel.Application") ' trap error 429
    If xlApp Is Nothing Then 'Excel not open
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    xlApp.DisplayAlerts = False 'do not display popups
    xlApp.Visible = blnxlAppVisible 'Show Excel App
    'Set xlWb = xlApp.Workbooks.Add
    'Call fct_TransferSpreadsheet("ContractList", strFilePathAndName)
    'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "ContractList", , True
    Set xlWb = xlApp.Workbooks.Add

    xlWb.SaveAs strFileName
    xlWb.Sheets("Sheet2").Delete
    xlWb.Sheets("Sheet3").Delete
    
    xlWb.Save
    xlWb.Close
    
    'xlApp.DisplayAlerts = True
  

    
lblCleanup:

  Set xlWb = Nothing
  Set xlApp = Nothing

fct_CreateEmptyExcelWorkbookSTD = True
    
Exit Function
  

Err_fct_ExcelCreateWorkbook:
    
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 429 ' Excel not open
            Resume Next
            ERR.Clear
        Case Else
            MsgBox "Error in fct_CreateEmptyExcelWorkbookSTD " & ERR.Number & ", Description: " & ERR.Description
            fct_CreateEmptyExcelWorkbookSTD = False
            Exit Function
        End Select
        GoTo lblCleanup
    End If



End Function

*****************
WORKSHEET - DELETE
*****************
        
        Public Function fct_ExcelDeleteWorksheetSTD(strFileName As String, strWsName As String) As Boolean
                       
          Dim xlApp As Excel.Application
          Dim xlWb As Excel.Workbook
          Dim xlWs As Excel.Worksheet
          
          On Error GoTo Err_fct_ExcelDeleteWorksheetSTD
          
            Set xlApp = GetObject(, "Excel.Application") ' must trap error 429 Excel not open
            If xlApp Is Nothing Then 'Excel not open
                Set xlApp = CreateObject("Excel.Application")
            End If
        
            Set xlWb = xlApp.Workbooks.Open(strFileName, , False)
        
            xlWb.Sheets(strWsName).Delete
        
            
            xlWb.Save
            xlWb.Close
            
            'xlApp.DisplayAlerts = True
          
        
            
        lblCleanup:
        
            Set xlWs = Nothing
          Set xlWb = Nothing
          Set xlApp = Nothing
        
        fct_ExcelDeleteWorksheetSTD = True
            
        Exit Function
          
        
        Err_fct_ExcelDeleteWorksheetSTD:
            
            If ERR.Number <> 0 Then
                Select Case ERR.Number
                Case 429 ' Excel not open
                    Resume Next
                    ERR.Clear
                Case Else
                    MsgBox "Error in fct_ExcelDeleteWorksheetSTD " & ERR.Number & ", Description: " & ERR.Description
                    fct_ExcelDeleteWorksheetSTD = False
                    Exit Function
                End Select
                GoTo lblCleanup
            End If
        
        
        
        End Function

*****************
TRANSFER SPREADSHEET
*****************
        
        Public Function fct_TransferSpreadsheetSTD(strQueryName As String, strFilePathAndName As String) As Boolean
        
        On Error GoTo Err_fct_TransferSpreadsheet
            Kill strFilePathAndName 'delete old file if one exists
        lblTransfer:
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strQueryName, strFilePathAndName, True
        
            fct_TransferSpreadsheetSTD = True
        
        Exit Function
        
        Err_fct_TransferSpreadsheet:
            
            If ERR.Number <> 0 Then
                Select Case ERR.Number
                Case 53 'kill file statement failed, continue
                    ERR.Clear
                    Resume Next
                Case 70 'Spreadsheet already exists and is open
                    MsgBox "Spreadsheet already exists and is open. Please close before proceeding."
                    ERR.Clear
                    GoTo lblTransfer
                Case 429 ' Excel not open
                    MsgBox "Excel not open"
                    Stop
                Case Else
                    MsgBox "Error in fct_TransferSpreadsheetSTD Nr: " & ERR.Number & ", Description: " & ERR.Description
                    fct_TransferSpreadsheetSTD = False
                    Exit Function
                End Select
                'GoTo lblcleanup
            End If
            
        End Function

*****************
WORKSHEETS - COPY
*****************
Public Function fct_CopyWorkbookSheetsAllSTD(strWBTemp As String, strWBTarget As String) As Boolean

  Dim xlApp As Excel.Application
  Dim xlWb, xlWBTarget, xlWBTemp As Excel.Workbook
  Dim xlWs As Excel.Worksheet
  
  
  On Error GoTo Err_fct_ExcelCopyWorkbook
   
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then 'Excel not open
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    ' Open original log (fPath = full path & filename for the orginal log)
    Set xlWBTarget = xlApp.Workbooks.Open(strWBTarget, , False)
    
    ' Open temporary log (xlTempLog = full path & filename for the temporary log)
    Set xlWBTemp = xlApp.Workbooks.Open(strWBTemp, , False)
    
    ' Copy any worksheet where the sheet name is a date
    For Each xlWs In xlWBTemp.Worksheets
         'If xlWs.Name = "ContractList" Then
              xlWs.Copy After:=xlWBTarget.Sheets(xlWBTemp.Sheets.Count)
              'xlWBTarget.Save
        ' End If
         
    Next xlWs
    
    xlWBTarget.Save
    xlWBTarget.Close
    xlWBTemp.Close

    Kill strWBTemp 'delete old file if one exists

lblCleanup:

  Set xlWs = Nothing
  Set xlWb = Nothing
  Set xlApp = Nothing

  
  fct_CopyWorkbookSheetsAllSTD = True
  
Exit Function

Err_fct_ExcelCopyWorkbook:
    
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 53 'file kill failed
        Resume Next
        Case 70 'kill failed, Workbook open, permission denied
            MsgBox "File open, permission denied, must close first. Error " & ERR.Number & ", Description: " & ERR.Description
            ERR.Clear
            Resume Next
        Case 91 'Object variable or With block variable not set
            Stop
        Case 429 ' Excel not open
            ERR.Clear
            Resume Next

        Case Else
            MsgBox "Error in fct_CopyWorkbookSTD Nr: " & ERR.Number & ", Description: " & ERR.Description
            fct_CopyWorkbookSheetsAllSTD = False
            Exit Function

        End Select
        GoTo lblCleanup
    End If



End Function

*****************
WORKSHEETS - HIDE
*****************
Public Function fct_HideWorksheets(strWB As String, Optional strWS As String) As Boolean

  Dim xlApp As Excel.Application
  Dim xlWb, xlWBTarget, xlWBTemp As Excel.Workbook
  Dim xlWs As Excel.Worksheet
  
  
  On Error GoTo Err_fct_ExcelCopyWorkbook
   
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then 'Excel not open
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    ' Open original log (fPath = full path & filename for the orginal log)
    Set xlWb = xlApp.Workbooks.Open(strWB, , False)
    
    
    ' Copy any worksheet where the sheet name is a date
    For Each xlWs In xlWb.Worksheets
         If xlWs.Name = strWS Or xlWs.Name Like "*(H)" Then
              xlWs.Visible = xlSheetHidden
              'xlWBTarget.Save
        End If
         
    Next xlWs
    
    xlWb.Save
    xlWb.Close

lblCleanup:

  Set xlWs = Nothing
  Set xlWb = Nothing
  Set xlApp = Nothing

  
  fct_HideWorksheets = True
  
Exit Function

Err_fct_ExcelCopyWorkbook:
    
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 53 'file kill failed
        Resume Next
        Case 70 'kill failed, Workbook open, permission denied
            MsgBox "File open, permission denied, must close first. Error " & ERR.Number & ", Description: " & ERR.Description
            ERR.Clear
            Resume Next
        Case 91 'Object variable or With block variable not set
            Stop
        Case 429 ' Excel not open
            ERR.Clear
            Resume Next

        Case Else
            MsgBox "Error in fct_HideWorksheets Nr: " & ERR.Number & ", Description: " & ERR.Description
            fct_HideWorksheets = False
            Exit Function

        End Select
        GoTo lblCleanup
    End If



End Function

*****************
FORMAT WORKBOOK
*****************

Public Function fct_ExcelFormatWorkbookSTD(strFilePathAndName As String, strWsName As String, Optional strWsNameNew As String) As Boolean
                
  Dim xlApp As Excel.Application

  Dim xlWb As Excel.Workbook
  Dim xlWs As Excel.Worksheet
  
  On Error GoTo Err_fct_ExcelFormatWorkbookSTD


    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then 'Excel not open
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Set xlWb = xlApp.Workbooks.Open(strFilePathAndName)
    Set xlWs = xlWb.Worksheets(strWsName)

    Call fct_ExcelFormatWorksheetSTD(xlWs, strWsNameNew)
    
    xlWb.Save
    xlWb.Close
    'xlApp.Quit

  
lblCleanup:

  Set xlWs = Nothing
  Set xlWb = Nothing
  Set xlApp = Nothing
  
fct_ExcelFormatWorkbookSTD = True
    
Exit Function
  
Err_fct_ExcelFormatWorkbookSTD:
    
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 429 ' Excel not open
            Resume Next
            ERR.Clear
        Case Else
            MsgBox "Error in fct_ExcelFormatWorkbookSTD, Error No: " & ERR.Number & ", Description: " & ERR.Description
            fct_ExcelFormatWorkbookSTD = False
            Exit Function
            
        End Select
        GoTo lblCleanup
    End If


  
End Function

*****************
FORMAT WORKSHEET
*****************

Function fct_ExcelFormatWorksheetSTD(xlWs, Optional strSheetNewName As String) As Boolean
    
On Error GoTo Err_fct_ExcelFormatWorksheet

Dim rgUsed As Range, rgTitleRow As Range
Dim FirstRow As Long, LastRow As Long, FirstCol As Long, LastCol As Long
Dim intCol As Integer
    
Set rgUsed = xlWs.UsedRange
    FirstRow = rgUsed(1).Row
    FirstCol = rgUsed(1).Column
    LastRow = rgUsed(rgUsed.Cells.Count).Row
    LastCol = rgUsed(rgUsed.Cells.Count).Column
    Set rgTitleRow = xlWs.Range("A1", xlWs.UsedRange.SpecialCells(xlLastCell))
    'Set rgTitleRow = rgUsed(Cells(1, 1), Cells(1, 1))
    
'GoTo lbltest
    rgUsed.AutoFilter
    rgUsed.Replace what:="&amp;", Replacement:="&"
    
    Call fct_ExcelColumnAutoWidthSTD(xlWs, 60)
    xlWs.Range("A2").Select
    xlWs.Application.ActiveWindow.FreezePanes = True
    xlWs.Range("A1", "A1").EntireRow.Font.Bold = True
    xlWs.Range("A1", "A1").EntireRow.Interior.Color = vbYellow

    'Format Amounts / Percent
    
    For intCol = 1 To LastCol ' this would be A to Z
    
      If InStr(1, xlWs.Cells(1, intCol), "Capital", 1) Or InStr(1, xlWs.Cells(1, intCol), "Amount", 1) Or InStr(1, xlWs.Cells(1, intCol), " Amt", 1) Or InStr(1, xlWs.Cells(1, intCol), "#", 1) Then
            xlWs.Cells(1, intCol).EntireColumn.NumberFormat = "#,##0.00"
      End If
      
      If InStr(1, xlWs.Cells(1, intCol), "Percent", 1) Or InStr(1, xlWs.Cells(1, intCol), "%", 1) Then
            xlWs.Cells(1, intCol).EntireColumn.NumberFormat = "0.00%"
      End If
      
      If InStr(1, Right(xlWs.Cells(1, intCol), 2), "ID", 1) Then
            xlWs.Cells(1, intCol).EntireColumn.Visible = False
      End If
  
      If InStr(1, Right(xlWs.Cells(1, intCol), 2), "(H)", 1) Then
            xlWs.Cells(1, intCol).EntireColumn.Visible = False
      End If
      
    Next intCol

'Dim intxlws.cellstatusIndex As Integer
    
'    intCellStatusIndex = xlWs.Range("A1:Z1").Find("Status").Column
    
'    Select Case xlWs.Name
'        Case "ContractList"
'            xlWs.Range("$A$1:$Y$114").AutoFilter Field:=intCellStatusIndex, Criteria1:="Active"
'        Case Else
'            Stop
'    End Select
    
    If strSheetNewName > "" Then xlWs.Name = strSheetNewName
    
lblCleanup:

  Set xlWs = Nothing
  'Set xlWB = Nothing
  'Set xlApp = Nothing
  
fct_ExcelFormatWorksheetSTD = True
    
Exit Function

Err_fct_ExcelFormatWorksheet:

    If ERR.Number <> 0 Then
        
        Select Case ERR.Number
        Case Else
            MsgBox "Error in fct_ExcelFormatWorksheetSTD, Error No: " & ERR.Number & ", Description: " & ERR.Description
            fct_ExcelFormatWorksheetSTD = False
            Exit Function
            'Stop
        End Select
        
    End If
    
End Function

*****************
EXCEL COLUMN AUTO WIDTH
*****************

Function fct_ExcelColumnAutoWidthSTD(xlWs, intMaxWidth As Integer) As Boolean

     Dim rgCell As Range
     'Application.ScreenUpdating = False
     
    For Each rgCell In xlWs.UsedRange.Rows(1).Cells
        rgCell.EntireColumn.AutoFit
        If rgCell.EntireColumn.ColumnWidth > intMaxWidth Then _
        rgCell.EntireColumn.ColumnWidth = intMaxWidth
     Next rgCell
     
    'Application.ScreenUpdating = True
    fct_ExcelColumnAutoWidthSTD = True
    
Exit Function
    
Err_fct_ExcelColumnAutoWidthSTD:

    If ERR.Number <> 0 Then
        
        Select Case ERR.Number
        Case Else
            MsgBox "Error in fct_ExcelColumnAutoWidthSTD, Error No: " & ERR.Number & ", Description: " & ERR.Description
            fct_ExcelColumnAutoWidthSTD = False
            Exit Function
            'Stop
        End Select
        
    End If
    
 End Function

*****************
EXCEL REFRESH ALL
*****************
Public Function fct_ExcelRefreshAll(strFileName As String, blnCloseXlApp As Boolean, blnShowRefreshPrompts As Boolean) As Boolean
               
    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlWs As Excel.Worksheet
    Dim xlAppWasOpen As Boolean
    Dim xlPivot As PivotTable
    
  xlAppWasOpen = True
  
  On Error GoTo Err_fct_ExcelRefreshAll
  
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then 'Excel not open
        Set xlApp = CreateObject("Excel.Application")
        xlAppWasOpen = False
    End If
    xlApp.DisplayAlerts = False 'alert that data may be in sheets, suppress
    xlApp.Visible = blnxlAppVisible
    Set xlWb = xlApp.Workbooks.Open(strFileName, , False)

    xlWb.RefreshAll
    DoEvents
    
    'refresh all does not refresh pivots!
    
        For Each xlWs In xlWb.Worksheets
            For Each xlPivot In xlWs.PivotTables
                        xlPivot.RefreshTable
                        'since only one pivot, we can exit
            Next xlPivot
        Next xlWs
    'xlWb.Save
    'Forms("0frm_Reports").SetFocus
    
    'msg to make sure refresh worked
    If blnShowRefreshPrompts Then MsgBox "Workbook data refresh, please wait until the last refresh date in Excel is current (do not close the workbook)"
    
    DoEvents
    xlWb.Save
    'Forms("0frm_Reports").SetFocus
    
    ' safety msg to make sure the refresh really worked
    If blnShowRefreshPrompts Then MsgBox "Workbook Refresh Complete"
    
    DoEvents
    xlWb.Close
    If xlAppWasOpen And blnCloseXlApp Then xlApp.Close ' set to false if multiple workbooks to refresh
    
    fct_ExcelRefreshAll = True
    
lblCleanup: 'executes in case of error too

  Set xlWb = Nothing
  Set xlApp = Nothing

Exit Function
  

Err_fct_ExcelRefreshAll:
    
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 429 ' Excel not open
            Resume Next
            ERR.Clear
        Case Else
            MsgBox "Error in fct_CreateEmptyExcelWorkbookSTD " & ERR.Number & ", Description: " & ERR.Description
            fct_ExcelRefreshAll = False
            Exit Function
        End Select
        GoTo lblCleanup
    End If

End Function

*****************
fct_ExcelRefreshAll_xlApp
*****************
Public Function fct_ExcelRefreshAll_xlApp(xlApp, strFileName As String) As Boolean
     
    'xlapp must be visible for refresh to work properly
    'Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlWs As Excel.Worksheet
    Dim xlPivot As PivotTable
    
  On Error GoTo Err_fct_ExcelRefreshAll_xlApp
  
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then 'Excel not open
        Set xlApp = CreateObject("Excel.Application")
        xlAppWasOpen = False
    End If
    
    xlApp.DisplayAlerts = False 'alert that data may be in sheets, suppress
    xlApp.Visible = True 'blnxlAppVisible
    
    Set xlWb = xlApp.Workbooks.Open(strFileName, , False)

    xlWb.RefreshAll
    DoEvents
    
    'refresh all does not refresh pivots!
    
        For Each xlWs In xlWb.Worksheets
            For Each xlPivot In xlWs.PivotTables
                        xlPivot.RefreshTable
                        'since only one pivot, we can exit
            Next xlPivot
        Next xlWs

    DoEvents
    xlWb.Save

    DoEvents
    xlWb.Close

    fct_ExcelRefreshAll_xlApp = True
    
lblCleanup: 'executes in case of error too

  Set xlWb = Nothing

Exit Function
  

Err_fct_ExcelRefreshAll_xlApp:
    
    If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 429 ' Excel not open
            Resume Next
            ERR.Clear
        Case Else
            MsgBox "Error in fct_CreateEmptyExcelWorkbookSTD " & ERR.Number & ", Description: " & ERR.Description
            fct_ExcelRefreshAll_xlApp = False
            Exit Function
        End Select
        GoTo lblCleanup
    End If



End Function
