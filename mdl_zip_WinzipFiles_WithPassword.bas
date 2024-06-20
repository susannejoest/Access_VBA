Public Function WinzipFilesWithPassword()

    Dim source As String
    Dim target As String
    Dim password As String
    
    source = Chr$(34) & "C:\DEPTS\050\" & Chr$(34)
    target = Chr$(34) & "C:\DEPTS\050\050" & Chr$(34)
    
    password = Chr$(34) & "JORDAN" & Chr$(34)
    Shell ("C:\Program Files\WinZip\WINZIP32.EXE -min -a -sPASSWORD " & target & " " & source)

End Function
