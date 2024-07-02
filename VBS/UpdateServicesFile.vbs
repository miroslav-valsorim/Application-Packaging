'==========================================================================
' Name: UpdateServices.vbs
' Version: 1.0
' Description: A script that can be used to update/append entries in the "C:\Windows\System32\drivers\etc\Services" file
'
' Usage: Add the desired new entries to a text file in the scriptdir. The default name of the file is "services-entries.txt". Can be changed in the sNewEntriesFile variable
'			The original services file is backed up with a filename defined in sBakFileName (e.g. "ServicesSAPInstall.bak")
'			Log is created in path location defined in sLogFolder
'			Important! In services-entries.txt on each line the first separator between the name and the port must be a tab symbol, e.g     "ssh	22/tcp". If needed any other spaces can be added after the tab
'==========================================================================

Option Explicit

On Error Resume Next

Dim oFS, oShell
Dim oLogFile
Dim sScriptDir, sLogFolder
Dim sBakFileName, sNewEntriesFile

Class HandleNewEntries
	Public NewLine, ExistFlag, UpdateFlag
	
	Public Property Get NewEntry
		If Len(NewLine) > 0 Then 
			NewEntry = Left(NewLine, InStr(1, NewLine, VBTab, vbTextCompare) - 1)
		Else
			NewEntry = Null
		End If
	End Property
End Class

Set oFS = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

sBakFileName = "ServicesSAPInstall.bak"
sLogFolder = "C:\Windows\Logs\"
sNewEntriesFile = "services-entries.txt"

CreateLogFile()
CheckServicesFile()


Set oFS = Nothing
Set oShell = Nothing

ExitScript(0)

Sub CheckServicesFile()
	Dim objTextFile
	Dim strServicesText
	Dim arrServicesText, arrServicesEntries, arrServicesNew()
	Dim strwinDir
	Dim index, i
	Dim bEntryFound

	Const ForReading = 1, ForWriting = 2, ForAppending = 8

	strwinDir = oShell.ExpandEnvironmentStrings("%WINDIR%")

	If (oFS.FileExists(sScriptDir & sNewEntriesFile)) Then
		If (oFS.FileExists(strwinDir & "\system32\drivers\etc\services")) Then
			WriteToLog "Creating backup of the Services file: " & strwinDir & "\system32\drivers\etc\" & sBakFileName
			oFS.CopyFile strwinDir & "\system32\drivers\etc\services", strwinDir & "\system32\drivers\etc\" & sBakFileName, 1
			
			' Read to array all new entries from sScriptDir & sNewEntriesFile
			Set objTextFile = oFS.OpenTextFile(sScriptDir & sNewEntriesFile, ForReading, True)
			If Not objTextFile.AtEndOfStream Then strServicesText = objTextFile.ReadAll()
			objTextFile.Close
			arrServicesText = Split(strServicesText, vbCrLf)

			For i = 0 To UBound(arrServicesText)
				ReDim Preserve arrServicesNew(i)
				Set arrServicesNew(i) = New HandleNewEntries
				arrServicesNew(i).NewLine = arrServicesText(i)
				arrServicesNew(i).ExistFlag = 0
				arrServicesNew(i).UpdateFlag = 0
			Next
			
			' Read to array all existing entries from strwinDir & "\system32\drivers\etc\services"
			Set objTextFile = oFS.OpenTextFile(strwinDir & "\system32\drivers\etc\services", ForReading, True)
			If Not objTextFile.AtEndOfStream Then strServicesText = objTextFile.ReadAll()
			objTextFile.Close
			arrServicesEntries = Split(strServicesText, vbCrLf)
			
			' Write/Update all entries that exist in services file. Append at the end the entries that were not found
			Set objTextFile = oFS.OpenTextFile(strwinDir & "\system32\drivers\etc\services", ForWriting, True)
		
			For i = 0 to UBound(arrServicesEntries)
				bEntryFound = 0
				For index = 0 to UBound(arrServicesNew)
					If InStr(1, arrServicesEntries(i), arrServicesNew(index).NewEntry, vbTextCompare) Then
						bEntryFound = 1
						arrServicesNew(index).ExistFlag = 1
						If Trim(arrServicesNew(index).NewLine) <> Trim(arrServicesEntries(i)) Then
							arrServicesNew(index).UpdateFlag = 1
							objTextFile.WriteLine Trim(arrServicesNew(index).NewLine)
							WriteToLog Chr(34) & Trim(arrServicesEntries(i)) & Chr(34) & " updated to " & Chr(34) & arrServicesNew(index).NewLine & Chr(34)
						Else
							objTextFile.WriteLine Trim(arrServicesEntries(i))
						End If
						Exit For
					End If
				Next
				If bEntryFound = 0 And Trim(arrServicesEntries(i)) <> "" Then objTextFile.WriteLine Trim(arrServicesEntries(i))
			Next

			For i = 0 to UBound(arrServicesNew)
				If arrServicesNew(i).ExistFlag = 0 And arrServicesNew(i).UpdateFlag = 0 And Trim(arrServicesNew(i).NewLine) <> "" Then
					objTextFile.WriteLine Trim(arrServicesNew(i).NewLine)
					WriteToLog Chr(34) & Trim(arrServicesNew(i).NewLine) & Chr(34) & " has been appended to the file"
				ElseIf arrServicesNew(i).ExistFlag = 1 And arrServicesNew(i).UpdateFlag = 0 And Trim(arrServicesNew(i).NewLine) <> "" Then
					WriteToLog Chr(34) & Trim(arrServicesNew(i).NewLine) & Chr(34) & " already exists in the file"
				End If
			Next

			objTextFile.Close	
		Else
			WriteToLog "Services file is missing in the machine... Aborting installation..."
			ExitScript(1)
		End If
	Else
		WriteToLog sNewEntriesFile & " is missing in the machine... Aborting installation..."
		ExitScript(1)
	End If

	Set objTextFile = Nothing
End Sub

' ===============================================================
' Write sExitText to log and quit, returning nExitCode
' ===============================================================
Sub ExitScript(nExitCode)
	Call WriteToLog("Script exit code: " & nExitCode)
	WriteToLog("------------- STOP LOGGING --------------")
	oLogFile.Close()
	WScript.Quit(nExitCode)
End Sub	'End ExitScript

' ========================================
' Create Log File
' ========================================
Sub CreateLogFile()
	' Dim sLogFileName : sLogFileName = Replace(WScript.ScriptFullName,".vbs",".log")
	Dim sLogFileName : sLogFileName = sLogFolder & Replace(WScript.ScriptName,".vbs",".log")
	
	'Backup log file if it has reached 5 MB size and create/append it
	If oFS.FileExists(sLogFileName) Then
		If oFS.GetFile(sLogFileName).Size > 5000000 Then
			oFS.CopyFile sLogFileName, sLogFileName & ".bak"
			Set oLogFile = oFS.OpenTextFile(sLogFileName, 2, True)
		Else
			Set oLogFile = oFS.OpenTextFile(sLogFileName, 8, True)
			oLogFile.WriteLine(vbCrLf)
		End If
	Else
		Set oLogFile = oFS.OpenTextFile(sLogFileName, 2, True)
	End If
	WriteToLog("------------- START LOGGING -------------")
	sScriptDir = left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))
	WriteToLog("ScriptDir: " & sScriptDir)
End Sub	'End CreateLogFile

' ========================================
' Write To Log
' ========================================
Sub WriteToLog(sTextToWrite)
	oLogFile.WriteLine(Now() & vbTab & sTextToWrite)
End Sub	'End WriteToLog