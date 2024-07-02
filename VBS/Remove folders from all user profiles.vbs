   '==========================================================================
'
' DESCRIPTION: Removes folders from all of the user profiles.
'
' NAME: HP_RemoveUserProfileFolder.vbs
'
' USAGE: Removes folder, along with subfolders and files.
'
' PREREQ: Ensure that the folder to removed does not contain files and 
'         folders from another application from the same suite.
'
' COMMENTS: None
'
'==========================================================================
Option Explicit

Const HKEY_LOCAL_MACHINE = &H80000002
Dim strComputer, objRegistry, strKeyPath, arrSubkeys, objSubkey, strValueName, strSubPath, strValue, strFolderName, strRemoveFolder

	'Name of the computer, "." determines the local computer
strComputer = "."

	'Name of the folder to be deleted under the user profiles i.e. BMC Software or BMC SOftware\Temp
strFolderName = "AppData\Roaming\Macromedia\Flash Player"
 
Set objRegistry=GetObject("winmgmts:\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
objRegistry.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
 
For Each objSubkey In arrSubkeys
    strValueName = "ProfileImagePath"
    strSubPath = strKeyPath & "\" & objSubkey
    objRegistry.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
    strRemoveFolder = strValue & "\" & strFolderName
    RemoveFolder strRemoveFolder
Next

Public Function RemoveFolder(strRemoveFolderName)
	Dim objFS, objShell, strSysDir
	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject( "WScript.Shell" )
	If (objFS.FolderExists(strRemoveFolderName)) Then
		objFS.DeleteFolder strRemoveFolderName, True
	End If
   	Set objShell = Nothing
   	Set objFS = Nothing
End Function