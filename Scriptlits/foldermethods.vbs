Set objFSO=CreateObject("Scripting.FileSystemObject")
Exportfile="c:\temp\McautoInstalls.csv"



'Sanity checks
if NOT objFSO.FolderExists("c:\temp\") Then
    objFSO.CreateFolder("c:\temp\")
END IF

if objFSO.FileExists(Exportfile) Then
    objFSO.DeleteFile(ExportFile)
END IF

'Create and Open The file to Write to
Set objFile = objFSO.CreateTextFile(Exportfile,True)
objfile.close