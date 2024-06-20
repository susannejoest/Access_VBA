Public Function fct_WinzipFilesWithPassword(strSource, strTarget, strPassword)

'Examples:
    'source = Chr$(34) & "C:\Temp\SourceFolder\" & Chr$(34)
    'target = Chr$(34) & "C:\Temp\TargetFolder" & Chr$(34)
    
    strPassword = Chr$(34) & "Password" & Chr$(34)
    Shell ("C:\Program Files\WinZip\WINZIP32.EXE -min -a -sPASSWORD " & target & " " & source)

End Function
