'==========================================================================
' Name: DeleteFilesAndFoldersFeomUserProfilesAndSystem.vbs
' Version: 1.0
' Description: Script to remove several files and/or folders from the users profile and from system, depending (or not depending) on whether they are empty or not.
'
' Usage: The folders to delete must be set in (see examples below):
'	aFoldersToDelete - Deletes the folder(s) from all user profiles even if they are not empty
'	aDeleteFolderIfEmpty - Deletes the folder(s) from all user profiles ONLY if they are empty
' 	aFoldersToDeleteFromSystem and aDeleteFolderFromSystemIfEmpty - have the same functionality as the above arrays but searches for folders in System
'	
'	The Files to delete must be set in:
'	aFilesToDelete - Deletes the file(s) from all user profiles
'	aFilesToDeleteFromSystem - Deletes the file(s) from System
'	Can be used as MSI custom action or run directly from command line. Doesn't create logs
'
'==========================================================================
Option Explicit

On Error Resume Next

Dim aFoldersToDelete, aDeleteFolderIfEmpty, aFoldersToDeleteFromSystem, aDeleteFolderFromSystemIfEmpty, aFilesToDelete, aFilesToDeleteFromSystem
Dim bDeleteIfEmpty
Dim bDeleteFiles
Dim sPFDir, sProgramDataDir
Dim oFS
Dim oShell

Set oFS = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
sPFDir = oShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
sProgramDataDir = oShell.ExpandEnvironmentStrings("%ProgramData%")

aFilesToDelete = Array("\AppData\Local\Launcher.config")
aFilesToDeleteFromSystem = Array(sPFDir & "\SAP\Common\saplogon.ini", sProgramDataDir & "\SAP\Common\saplogon.ini", sPFDir & "\SAP\Common\SAPUILandscape.xml", sProgramDataDir & "\SAP\Common\SAPUILandscape.xml")

'aFoldersToDelete = Array("\AppData\Roaming\Pulse Secure\Setup Client")
'aDeleteFolderIfEmpty = Array("\AppData\Local\Plantronics", "\AppData\Local\Sennheiser")

'aFoldersToDeleteFromSystem = Array(sPFDir & "\OpenScape Desktop Client", sProgramDataDir & "\Sennheiser\SDKCore")
'aDeleteFolderFromSystemIfEmpty = Array(sProgramDataDir & "\Sennheiser")

'Delete Files from User Profiles
bDeleteFiles = 0
If IsBound(aFilesToDelete) Then ScanUsers(aFilesToDelete)
If IsBound(aFilesToDeleteFromSystem) Then DeleteFromSystem(aFilesToDeleteFromSystem)
bDeleteFiles = 1

'Delete folders from User profiles
bDeleteIfEmpty = 0
If IsBound(aFoldersToDelete) Then ScanUsers(aFoldersToDelete)
bDeleteIfEmpty = 1
If IsBound(aDeleteFolderIfEmpty) Then ScanUsers(aDeleteFolderIfEmpty)

'Delete folders from System
bDeleteIfEmpty = 0
If IsBound(aFoldersToDeleteFromSystem) Then DeleteFromSystem(aFoldersToDeleteFromSystem)
bDeleteIfEmpty = 1
If IsBound(aDeleteFolderFromSystemIfEmpty) Then DeleteFromSystem(aDeleteFolderFromSystemIfEmpty)

' ========================================
' ScanUsers
' ========================================
Sub ScanUsers(aFolders)
	On Error Resume Next
	Dim sUsersDir
	Dim oBaseFolder, colFolders, oFolders, i
	
	sUsersDir = oShell.ExpandEnvironmentStrings("%SystemDrive%\Users\")
	If oFS.FolderExists(sUsersDir) Then
		Set oBaseFolder = oFS.GetFolder(sUsersDir)
		Set colFolders = oBaseFolder.Subfolders
		For Each oFolders In colFolders
			If right(oFolders, 9) <> "All Users" AND right(oFolders, 12) <> "Default User" AND right(oFolders, 6) <> "Public" Then			
				For i = 0 To UBound(aFolders)
					If bDeleteFiles Then
						If oFS.FileExists(oFolders & aFolders(i)) Then oFS.DeleteFile oFolders & aFolders(i), True
					Else
						If oFS.FolderExists(oFolders & aFolders(i)) Then
							If bDeleteIfEmpty Then
								If FolderEmpty(oFolders & aFolders(i)) Then	oFS.DeleteFolder oFolders & aFolders(i), True
							Else
								oFS.DeleteFolder oFolders & aFolders(i), True
							End If
						End If
					End If
				Next
			End If
		Next
	End If
End Sub

' ========================================
' DeleteFromSystem
' ========================================
Sub DeleteFromSystem(aFolders)
	On Error Resume Next
	Dim sFolderName
	
	For Each sFolderName In aFolders
		If bDeleteFiles Then
			If oFS.FileExists(sFolderName) Then oFS.DeleteFile sFolderName, True
		Else
			If oFS.FolderExists(sFolderName) Then
				If bDeleteIfEmpty Then
					If FolderEmpty(sFolderName) Then oFS.DeleteFolder sFolderName, True
				Else
					oFS.DeleteFolder sFolderName, True
				End If
			End If
		End If
	Next
End Sub

' ========================================
' IsBound
' ========================================
Function IsBound(inArray)
	On Error Resume Next

	If IsArray(inArray) And UBound(inArray) >= 0 Then
		IsBound = True
	Else
		IsBound = False
	End If
	If Err.Number <> 0 Then
		Err.Clear
		IsBound = False
	End If
End Function

' ========================================
' FolderEmpty
' ========================================
Function FolderEmpty(sFolder)
	Dim oFolder
	
	If oFS.FolderExists(sFolder) Then
		Set oFolder = oFS.GetFolder(sFolder)
  
		If oFolder.Files.Count = 0 And oFolder.SubFolders.Count = 0 Then
			FolderEmpty=True
		Else
			FolderEmpty=False
		End If
	End If
End Function