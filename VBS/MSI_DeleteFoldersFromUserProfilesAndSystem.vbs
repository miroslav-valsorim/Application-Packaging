'==========================================================================
' Name: DeleteFoldersFromUserProfilesAndSystem.vbs
' Version: 1.1
' Description: Script to remove several folders from the users profile and from system, depending (or not depending) on whether they are empty or not.
'
' Usage: The folders to delete must be set in (see examples below):
'	aFoldersToDelete - Deletes the folder(s) from all user profiles even if they are not empty
'	aDeleteFolderIfEmpty - Deletes the folder(s) from all user profiles ONLY if they are empty
' 	aFoldersToDeleteFromSystem and aDeleteFolderFromSystemIfEmpty - have the same functionality as the above arrays but searches for folders in System
'	Used as MSI custom action. Writes in the msi log
'
'==========================================================================
Option Explicit

On Error Resume Next

Dim aFoldersToDelete, aDeleteFolderIfEmpty, aFoldersToDeleteFromSystem, aDeleteFolderFromSystemIfEmpty
Dim bDeleteIfEmpty
Dim sPFDir, sProgramDataDir
Dim oFS
Dim oShell
Dim sFunctionString

Set oFS = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
sPFDir = oShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
sProgramDataDir = oShell.ExpandEnvironmentStrings("%ProgramData%")

aFoldersToDelete = Array("\AppData\LocalLow\Sun\Java", "\AppData\Roaming\Sun\Java", "\AppData\Roaming\Oracle\Java")
aDeleteFolderIfEmpty = Array("\AppData\LocalLow\Sun", "\AppData\Roaming\Sun", "\AppData\Roaming\Oracle")
aFoldersToDeleteFromSystem = Array(sProgramDataDir & "\Microsoft\Windows\Start Menu\Programs\Java")

sFunctionString = "DeleteFiles"
Call WriteToLog("####################################################################################################################")

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

Call WriteToLog("####################################################################################################################")

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
					If oFS.FolderExists(oFolders & aFolders(i)) Then
						If bDeleteIfEmpty Then
							If FolderEmpty(oFolders & aFolders(i)) Then
								Call WriteToLog(vbTab & "Delete " & oFolders & aFolders(i))
								oFS.DeleteFolder oFolders & aFolders(i), True
							Else
								Call WriteToLog(vbTab & oFolders & aFolders(i) & " is not empty")
							End If
						Else
							Call WriteToLog(vbTab & "Delete " & oFolders & aFolders(i))
							oFS.DeleteFolder oFolders & aFolders(i), True
						End If
					Else
						Call WriteToLog(vbTab & oFolders & aFolders(i) & " doesn't exist")
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
		If oFS.FolderExists(sFolderName) Then
			If bDeleteIfEmpty Then
				If FolderEmpty(sFolderName) Then
					Call WriteToLog(vbTab & "Delete " & sFolderName)
					oFS.DeleteFolder sFolderName, True
				Else
					Call WriteToLog(vbTab & sFolderName & " is not empty")
				End If
			Else
				Call WriteToLog(vbTab & "Delete " & sFolderName)
				oFS.DeleteFolder sFolderName, True
			End If
		Else
			Call WriteToLog(vbTab & sFolderName & " doesn't exist")
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

' ========================================
' Write to LOG file
' ========================================
Sub WriteToLog(sTextToWrite)
	Const msiMessageTypeInfo = &H04000000
	Dim record
	
	If sFunctionString = "" Then sFunctionString = "CustomActionScriptLog"

    Set record = Installer.CreateRecord(1)
    record.stringdata(0) = sFunctionString & ": [1]"
    record.stringdata(1) = sTextToWrite
    record.formattext
    message msiMessageTypeInfo, record
    Set record = Nothing
End Sub
