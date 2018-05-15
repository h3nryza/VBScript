On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if

'Declarations
Const DeleteReadOnly=TRUE
Dim objFSO,wshShell
Dim objF,arrFiles,objfl

'Sets
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject( "WScript.Shell" )



Sub Subfolder (objFolder)
    'Check all files within this directory
    getListofFiles objFolder
    
    'Get All Subfolders within the Folder
    For Each objSubFolder in objFolder.Subfolders
        'Call sub to get all files within the subfolder
        getListofFiles objSubFolder
        
        'Delete Folder
        objFso.deletefolder objSubFolder,True
        'Loop back in to get subfolders within the subfolder
        SubFolder objSubFolder
    
    Next
End Sub




Sub getListofFiles (objFolder)     
    
    'Set folder
    set objF = objfso.getfolder(objFolder)
    'List array of files
    set arrFiles = objF.files
    

    'Max number of items folder can contain
    cntMaxFlderItems = 100

    'If the Max amount of files in the folder has ben reached, Send the files for deletion
    if arrFiles.length > cntMaxFlderItems Then
        'Check if the files are ready for deletion
        For Each objfile in arrFiles
            delFiles objfile
        Next
    End If

End Sub




Sub delFiles(objfile)
    
    'Set File
    set objfl = objFSO.GetFile(objfile)
    'Set File Date
    objFileDate = objfl.DateLastModified
    'Set Date to compate against
    dtCompare = dateadd("d",-7,now())
    
    'Compare to see if file is older than the comparision date, delete if true
    If objFileDate <  dtCompare Then
        objfso.deletefile objfile
  
    End If

End Sub



Subfolder objfso.getfolder("C:\temp")
