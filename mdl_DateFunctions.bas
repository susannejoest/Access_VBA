Attribute VB_Name = "mdl_DateFunctions"
Option Compare Database

Function GetFormattedDate(strFormat As String, strMonth As String, Optional strDay As String) As String

' strFormat "yyyymmdd" "yymm" "mmyy", Month = "CurrentMonth" "PreviousMonth", Day = "FirstDay" (default if null) "LastDay"

    Dim dtDate As Date
    Dim dtDay As Date
    
    ' Get the current date
    Select Case strMonth
    
        Case "CurrentMonth"
            dtDate = Date
        Case "PreviousMonth"
            dtDate = DateAdd("m", -1, Date)
        Case Else
            Stop
    End Select
    
    Select Case strDay
        Case Null ' First day is default
            dtDay = DateSerial(Year(dtDate), Month(dtDate), 1)
        Case "FirstDay"
            dtDay = DateSerial(Year(dtDate), Month(dtDate), 1)
        Case "LastDay"
            dtDay = DateSerial(Year(dtDate), Month(dtDate) + 1, 0)
        Case Else
            Stop
    End Select
    

    ' Return the last day of the previous month in yyyymmdd format
    Select Case strFormat

        Case "yyyymmdd"
            GetFormattedDate = Format(strDay, "yyyymmdd")
        Case "yymm"
            GetFormattedDate = Format(strDay, "yymm")
        Case "mmyy"
            GetFormattedDate = Format(strDay, "mmyy")
            
    End Select
    
End Function
