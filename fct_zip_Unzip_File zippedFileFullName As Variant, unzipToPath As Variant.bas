Function UnzipAFile(zippedFileFullName As Variant, unzipToPath As Variant)
    
    Dim ShellApp As Object
    
    'Copy the files & folders from the zip into a folder
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(unzipToPath).CopyHere ShellApp.Namespace(zippedFileFullName).Items

End Function
