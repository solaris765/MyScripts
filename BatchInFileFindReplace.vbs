Option Explicit

Dim strPath, objFSO, objFolder, colFiles, objFile, varFind, varReplace, extension

'Brings up a Directory Selector and asks User to choose one
strPath = SelectFolder( "" )
If strPath = vbNull Then
    WScript.Echo "Cancelled"
Else
    WScript.Echo "Selected Folder: """ & strPath & """"
End If

'Object Definitions
Set objFSO = CreateObject("scripting.filesystemobject")
Set objFolder = objFSO.GetFolder(strPath)
Set colFiles = objFolder.Files

'User Input
varFind = inputBox("Find text to replace:", "Find")
varReplace = inputBox("What would you like to replace it with?:", "Replace")

'Extension User Input and assumptions
extension = ""
extension = inputbox("What extension is the file type? no '.':", "Extension")
If extension = "" then
 extension = "txt"
End If

'For loop to iterate through all files in root folder
For Each objFile in colFiles
 If objfso.GetExtensionName(objFile) = extension then
  Call inFileFindReplace((strPath + "\" + objFile.name), varFind, varReplace)
 End If
Next


' In file find and replace function
Function inFileFindReplace(inFile, strFind, strReplace)
 Dim strText, strNewText
 Const ForReading = 1

 Const ForWriting = 2

 Set objFSO = CreateObject("Scripting.FileSystemObject")

 'For Reading Object
 Set objFile = objFSO.OpenTextFile(inFile, ForReading)

'read from file
 strText = objFile.ReadAll
'close file
 objFile.Close
'replace text in file 
 strNewText = Replace(strText, strFind, strReplace)

'create outfile obj
 Set objFile = objFSO.OpenTextFile(inFile, ForWriting)
'write changes to file
 objFile.WriteLine strNewText
'close file
 objFile.Close
End Function

'Select Folder function by Rob van der Woude
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