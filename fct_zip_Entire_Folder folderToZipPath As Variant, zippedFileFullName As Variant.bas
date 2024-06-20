Function CreateZipFile_ENTIREFOLDER(folderToZipPath As Variant, zippedFileFullName As Variant)
    ' usage: Call CreateZipFile("C:\Users\marks\Documents\ZipThisFolder\", "C:\Users\marks\Documents\NameOFZip.zip")
    
    Dim ShellApp As Object

    'Create an empty zip file
    Open zippedFileFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    'Copy the files & folders into the zip file
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(zippedFileFullName).CopyHere ShellApp.Namespace(folderToZipPath).Items
    
    'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
    On Error Resume Next
    Do Until ShellApp.Namespace(zippedFileFullName).Items.Count = ShellApp.Namespace(folderToZipPath).Items.Count
    Stop
        'Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0

End Function
