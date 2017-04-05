Option Explicit
Dim root, FindStr, ReplaceStr

root = SelectFolder( "" )
If root = vbNull Then
    WScript.Echo "Cancelled"
'Else
    'WScript.Echo "Selected Folder: """ & root & """"
End If

FindStr = INPUTBOX("Please enter the string that you want to find")
IF FindStr = "" THEN Canceled

ReplaceStr = INPUTBOX("Please enter the string that you want to replace it with")
'IF ReplaceStr = "" THEN Canceled

RenameFolderRecurse root, FindStr, ReplaceStr


Sub RenameFolderRecurse (folder, strFind, strReplace)
 Dim objFSO, name, objFolder, sf, newName
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 
 Set objFolder = objFSO.GetFolder(folder)
 
 for each sf In objFolder.SubFolders
  RenameFolderRecurse sf, strFind, strReplace
 Next
 
 name = objFSO.GetFolder(folder).Name

 newName = Replace(name, strFind, strReplace)

 If name <> newName Then
  If objFSO.FolderExists(folder) Then
   objFSO.GetFolder(folder).Name = newName
  End If
 End If
 
End Sub

Function SelectFolder( myStartFolder )
 ' This function opens a "Select Folder" dialog and will
 ' return the fully qualified path of the selected folder
 '
 ' Argument:
 '     myStartFolder    [string]    the root folder where you can start browsing;
 '                                  if an empty string is used, browsing starts
 '                                  on the local computer
 '
 ' Returns:
 ' A string containing the fully qualified path of the selected folder
 '
 ' Written by Rob van der Woude
 ' http://www.robvanderwoude.com
 
    ' Standard housekeeping
    Dim objFolder, objItem, objShell
    
    ' Custom error handling
    On Error Resume Next
    SelectFolder = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, myStartFolder )

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function