'==========================================================================
' Script: UpdateTNSNames.vbs
' Version: 1.0
' Description: The script updates tnsnames.ora with desired value(s). The algorithm queries one or more tnsnames.ora locations. 
' 	- By default it's looking in the location from %TNS_ADMIN% and in the Oracle client admin folder (set by sOracleClientDir variable).
' 	- If a location doesn't exist, the process is skipped.
' 	- If the folder exists, but there is no tnsnames.ora in it, then the file is created with default value set in sNewOra variable. A header is created (defined by sUpdateHeader & sEndOfHeader)
'	- If the file exists then its queried for sLookForEntry. If the entry doesn't exist then sNewOra is added to the end of the file. If the entry is found then it's updated to
'		 sNewHostName (if needed). If any updates are made, header is also updated with sUpdateHeader
'
' USAGE: Set folders to query with different variables. Default folders are set in sTNSNamesDir and sOracleClientDir. Additional folders can be added. Each folder is queried by QueryFolder() function
'			Set the entry to look for in sLookForEntry. Set the name of the server in sNewHostName. The string that is used to update header can be set in sUpdateHeader.
			'sNewOra is the string that will be written if sLookForEntry is not found in the file or if the file is missing, but the folder exists.
'
' CHANGELOG: 13.04.2020 - Initial version
'==========================================================================

Option Explicit

' On Error Resume Next

Dim oFS, oShell
Dim oLogFile
Dim sScriptDir, sLogFolder
Dim sBakFileName
Dim sTNSNamesDir, sOracleClientDir, sCurrentWorkDir
Dim sLookForEntry, sNewHostName, sEndOfHeader, sUpdateHeader, sNewOra

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set oFS = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

sLogFolder = "C:\Windows\Logs\"
sTNSNamesDir = oShell.ExpandEnvironmentStrings("%TNS_ADMIN%\")
sOracleClientDir = "C:\Oracle\product_32Bit\12.1.0\client_1\network\admin\"
sBakFileName = "TNSNames.ora_VISIONUpdate.bak"
sLookForEntry = RemoveWhiteSpaces("visionp.world=")
sNewHostName = "njazgvpvis01"
sEndOfHeader = "#################################################################"
sUpdateHeader = "# " & Right("0" & DatePart("d",Date) & "/", 3) & Right("0" & DatePart("m",Date) & "/", 3) & DatePart("yyyy",Date) _
				& " DXC - VISIONP host updated to njazgvpvis01 by Vision package" & vbCrLf & sEndOfHeader

sNewOra = "visionp.world =  " & vbCrLf & _
"  (DESCRIPTION =  " & vbCrLf & _
"        (ADDRESS =  " & vbCrLf & _
"          (COMMUNITY = TCP.world) " & vbCrLf & _
"          (PROTOCOL = TCP) " & vbCrLf & _
"          (Host = njazgvpvis01)" & vbCrLf & _
"          (Port = 1521) " & vbCrLf & _
"        ) " & vbCrLf & _
"    (CONNECT_DATA =  " & vbCrLf & _
"      (SID = visionp)" & vbCrLf & _
"         (GLOBAL_NAME = visionp.world) " & vbCrLf & _
"    ) " & vbCrLf & _
"  )" & vbCrLf


CreateLogFile()

QueryFolder(sTNSNamesDir)
QueryFolder(sOracleClientDir)

ExitScript(0)

' ===============================================================
' QueryFolder
' ===============================================================
Sub QueryFolder(sFolderToCheck)
	If oFS.FolderExists(sFolderToCheck) Then
		sCurrentWorkDir = sFolderToCheck ' sCurrentWorkDir will be used in case of failure to restore the original ora file
		UpdateTNSNAMES(sFolderToCheck)
	Else
		WriteToLog sFolderToCheck & " doesn't exist. Skipping TNSNames.ora update for " & sFolderToCheck
	End If
End Sub	'End QueryFolder

' ===============================================================
' UpdateTNSNAMES
' ===============================================================
Sub UpdateTNSNAMES(sORALocation)
	Dim bEntryFound: bEntryFound = 0
	Dim bFileUpdated: bFileUpdated = 0
	Dim oTNSNAMESOLD, oTNSNAMES
	Dim sLine, sFoundHostName
	Dim i
	
	If (oFS.FileExists(sORALocation & "TNSNames.ora")) Then
		WriteToLog "Creating backup of the TNSNames.ora file: " & sORALocation & sBakFileName
		oFS.CopyFile sORALocation & "TNSNames.ora", sORALocation & sBakFileName, 1
		
		Set oTNSNAMESOLD = oFS.OpenTextFile(sORALocation & sBakFileName, ForReading, True)
		Set oTNSNAMES = oFS.OpenTextFile(sORALocation & "TNSNames.ora", ForWriting, True)
		
		WriteToLog "Updating TNSNames.ora file in " & sORALocation
		Do Until oTNSNAMESOLD.AtEndOfStream
			sLine = oTNSNAMESOLD.Readline
			If InStr(1, Left(RemoveWhiteSpaces(sLine), Len(sLookForEntry)), sLookForEntry, vbTextCompare) > 0 Then
				WriteToLog sLookForEntry & " found"
				oTNSNAMES.WriteLine sLine
				bEntryFound = 1
				i = 1
				Do While 1
					sLine = oTNSNAMESOLD.Readline
					If InStr(1, sLine, "(Host", vbTextCompare) > 0 Then Exit Do
					oTNSNAMES.WriteLine sLine
					i = i + 1
					If i > 9 Then
						WriteToLog "Something went wrong! No host information found for " & sLookForEntry & "! TNSNames.ora will be aborted!"
						oTNSNAMESOLD.Close
						oTNSNAMES.Close
						ExitScript(1603)
					End If
				Loop
				
				sFoundHostName = FindHostName(sLine)
				If InStr(1, sFoundHostName, sNewHostName, vbTextCompare) > 0 Then
					WriteToLog "Host is already set to " & sNewHostName
					oTNSNAMES.WriteLine sLine
					bFileUpdated = 0
				Else
					WriteToLog "Updating Host to " & sNewHostName
					oTNSNAMES.WriteLine Replace(sLine, sFoundHostName, sNewHostName)
					bFileUpdated = 1
				End If
			Else
				If Not oTNSNAMESOLD.AtEndOfStream Then
					oTNSNAMES.WriteLine sLine
				Else
					oTNSNAMES.Write sLine
				End If
			End If
		Loop
		
		oTNSNAMESOLD.Close
		oTNSNAMES.Close
		
		If bEntryFound = 0 Then
			' The entry wasn't found. Append it to end of file
			WriteToLog sLookForEntry & " wasn't found. Append it to end of file"
			Set oTNSNAMES = oFS.OpenTextFile(sORALocation & "TNSNames.ora", ForAppending, True)
			oTNSNAMES.Write sNewOra
			oTNSNAMES.Close
			bFileUpdated = 1
		End If
	Else
		' "TNSNames.ora file is missing in " & sORALocation & ". New file will be created."
		WriteToLog "TNSNames.ora file is missing in " & sORALocation & ". New file will be created."
		Set oTNSNAMES = oFS.OpenTextFile(sORALocation & "TNSNames.ora", ForWriting, True)
		oTNSNAMES.Write "# uses host namesx" & vbCrLf & sUpdateHeader & vbCrLf
		oTNSNAMES.Write sNewOra
		oTNSNAMES.Close
	End If
	
	' Update file header
	If bFileUpdated = 1 Then
		Set oTNSNAMES = oFS.OpenTextFile(sORALocation & "TNSNames.ora", ForReading, True)
		Dim sUpdateText: sUpdateText = oTNSNAMES.ReadAll()
		oTNSNAMES.Close
		sUpdateText = Replace(sUpdateText, sEndOfHeader, sUpdateHeader, 1, 1)
		Set oTNSNAMES = oFS.OpenTextFile(sORALocation & "TNSNames.ora", ForWriting, True)
		oTNSNAMES.Write sUpdateText
		oTNSNAMES.Close
	End If
End Sub	'End UpdateTNSNAMES

' ===============================================================
' FindHostName
' ===============================================================
Function FindHostName(sLineWithHostName)
	Dim iBegin, iEnd
	iBegin = InStr(1, RemoveWhiteSpaces(sLineWithHostName), "(Host=", vbTextCompare)
	iEnd = InStr(iBegin, RemoveWhiteSpaces(sLineWithHostName), ")", vbTextCompare)
	FindHostName = Mid(RemoveWhiteSpaces(sLineWithHostName), iBegin + Len("(Host="), iEnd - iBegin - Len("(Host="))
End Function	'End FindHostName

' ===============================================================
' RemoveWhiteSpaces
' ===============================================================
Function RemoveWhiteSpaces(sWithWhiteSpaces)
	Dim RegX
	Set RegX = New RegExp
	
	RegX.Pattern = "\s"
	RegX.Global = True
	RemoveWhiteSpaces = RegX.Replace(sWithWhiteSpaces, "")
End Function	'End RemoveWhiteSpaces

' ===============================================================
' Write sExitText to log and quit, returning nExitCode
' ===============================================================
Sub ExitScript(nExitCode)
	If nExitCode = 1603 Then
		Call WriteToLog("Restoring original TNSNames.ora")
		oFS.CopyFile sCurrentWorkDir & sBakFileName, sCurrentWorkDir & "TNSNames.ora", 1
	End If
	Call WriteToLog("Script exit code: " & nExitCode)
	WriteToLog("------------- STOP LOGGING --------------")
	oLogFile.Close()
	Set oLogFile = Nothing
	Set oFS = Nothing
	Set oShell = Nothing
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
			Set oLogFile = oFS.OpenTextFile(sLogFileName, 2, True)
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