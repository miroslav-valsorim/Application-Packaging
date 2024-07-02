'************************************************************************************************************************
'*		Script: DefaultUser.vbs																							*
'*		Version: 1.0.5																									*													*
'*																														*
'*		Usage: 	This is a modification of the DefaultUser.vbs, that can be executed only as	custom action from MSI		*
'*			The input file is set in strInputRegFileName																*
'*			The script assumes the input file is deployed in the folder of the msi, defined as the CACHEDIR property in *
'			the msi																										*
'*			This file will then be parsed where all the HKEY_CURRENT_USER settings will be changed						*
'*			to point to a DEFAULT USER registry key (which can be changed using the USER_PROFILE_REG					*
'*			constant).																									*
'*			Returns standard MSI CA return codes: Success=1; Error=3													*
'*			VBS-Stored in the Binary Table. Script Function: ApplyRegistryToProfiles. Condition: as required.			*
'*			Warning!!! Session.Property is used, and since most of the properties are not visible in Deferred Execution,*
'*			Immediate Execution must be used for the CA, or alternatively the according property must be made available *
'*			in Deferred phase																							*
'*																														*
'*																														*
'*		V1.0.1:	Changed enumeration of all users in the WMI, to an enumeration of the ProfilesList.						*
'*		v1.0.2: Added the Loadhive for the Default User. This user was not handled.										*
'*		v1.0.3:	Added enumeration of HKEY_USERS to keep track of hives that are already loaded 							*
'*			(user is no longer required to log off).																	*
'*		v1.0.4: Changed behaviour to look up the Default User's path, as this has changed with Vista.					*
'*		v1.0.5: Customized to be called from MSI Custom Action															*
'*																														*
'************************************************************************************************************************

Option Explicit

CONST   CSTR_ForReading 	= 1
CONST   CSTR_ForWriting 	= 2
CONST   CSTR_ForAppending 	= 8
CONST	HKEY_LOCAL_MACHINE 	= &H80000002
CONST	HKEY_USERS	 	= &H80000003

CONST	DEFAULT_USER_FILE	= "DefUser_"
CONST	CURRENT_USER_REG	= "HKEY_CURRENT_USER"
CONST	USERS_REG		= "HKEY_USERS"
CONST	USER_PROFILE_REG	= "HKEY_LOCAL_MACHINE\Temp"
CONST	PROFILE_REG_BASE	= "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
CONST	WINDOWSNT_VERSION_REG	= "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

Private oFSO, oWshShell, oWshPrsEnv, oWMI, oReg
Private strInputRegFile, strInputRegFileName, strOutputRegFile, strFileName
Private arrSplitFilePath
Private intReturn
Private sFunctionString, strMSIPath

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWshShell = CreateObject("WScript.Shell")
Set oWshPrsEnv = oWshShell.Environment("PROCESS")
Set oReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!//" & oWshPrsEnv("ComputerName") & "/root/default:StdRegProv")

'Set InputRegFileName
strInputRegFileName = "remove.reg"

Function ApplyRegistryToProfiles()
	On Error Resume Next

	Dim arrProfiles
	Dim objProfile
	Dim strProfileImagePath
	Dim strDefaultUserProfile, strDefaultUserPath, strWindowsVersion
	Dim arrUserProfileKey
	Dim objUserProfileKey
	Dim arrUserProfile
	
	sFunctionString = "ApplyRegistryToProfiles"
	ApplyRegistryToProfiles = 1
	Call Report("####################################################################################################################")
	Call Report("Script start ...")
	
	strMSIPath = Session.Property("CACHEDIR")
	If Right(strMSIPath, 1) <> "\" Then strMSIPath = strMSIPath & "\"
	strInputRegFile = strMSIPath & strInputRegFileName

	Call Report(vbTab & "> Using " & strInputRegFile & " as input file.")
	Call Report(vbTab & "# Checking file existence.")
	If oFSO.FileExists(strInputRegFile) Then
		strFileName = DEFAULT_USER_FILE & strInputRegFileName
		If Right(strFileName, 1) = """" Then
			strFileName = Mid(strFileName, 1, Len(strFileName) -1)
		End If
		If oWshPrsEnv("TEMP") <> vbNullString Then
			strOutputRegFile = oWshPrsEnv("TEMP") & "\"
		Else
			strOutputRegFile = strMSIPath
		End If
		strOutputRegFile = strOutputRegFile & strFileName
		Call Report(vbTab & "> Inputfile found. Using " & strOutputRegFile & " as output file.")
	Else
		ApplyRegistryToProfiles = 3
		Call ExitScript("Function exit code: " & ApplyRegistryToProfiles)
		Exit Function
	End If
	' Call Report("# Changing file from CURRENT_USER to DEFAULT_USER")
	' ChangeRegFile CURRENT_USER_REG, USER_PROFILE_REG

	Call Report(vbTab & "# Applying registry to Profiles List.")

	' First we load the hive from the Default User and change the corresponding registry settings.
	Call Report(vbTab & vbTab & "> Changing file from CURRENT_USER to " & USER_PROFILE_REG)
	ChangeRegFile CURRENT_USER_REG, USER_PROFILE_REG
	Call Report(vbTab & "> Pushing settings for the Default User")
	oReg.GetStringValue HKEY_LOCAL_MACHINE, WINDOWSNT_VERSION_REG, "CurrentVersion", strWindowsVersion
	If Left(strWindowsVersion, 1) > 5 Then
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, "Default", strProfileImagePath
	Else
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, "ProfilesDirectory", strDefaultUserPath
		oReg.GetStringValue HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, "DefaultUserProfile", strDefaultUserProfile
		strProfileImagePath = strDefaultUserPath & "\" & strDefaultUserProfile
	End If
	Call Report(vbTab & vbTab & vbTab & "> Loading Registry Hive into " & USER_PROFILE_REG)
	If LoadHive(strProfileImagePath) = 0 Then
		Call Report(vbTab & vbTab & vbTab & "> Adding registry settings from file " & strOutputRegFile)
		intReturn = oWshShell.Run("Regedit /s """ & strOutputRegFile & """", 0, True)
		Call Report(vbTab & vbTab & vbTab & "> Unloading User Registry Hive from " & USER_PROFILE_REG)
		UnloadHive
	Else
		Call Report(vbTab & vbTab & vbTab & "> Could not Load Hive for the Default User(another hive may already be loaded)! Skipping...")
	End If

	' Next, all profiles are enumerated
	Call Report(vbTab & "> Enumerating Profiles")
	oReg.EnumKey HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, arrProfiles
	If IsBound(arrProfiles) Then
		For Each objProfile in arrProfiles
			If Err.Number <> 0 Then
				ApplyRegistryToProfiles = 3
				Call ExitScript(vbTab & "> " & Err.Description)
				Exit Function
			End If
			' UserProfiles start with S-1-5-21-... All other profiles, like Network and System, are ignored.
			If Instr(1, objProfile, "S-1-5-21-", vbTextCompare) Then
				Call Report(vbTab & vbTab & "# Checking Profile: " & objProfile)
				oReg.EnumKey HKEY_USERS, objProfile, arrUserProfileKey
				' Check whether the profile is already loaded.
				oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, PROFILE_REG_BASE & "\" & objProfile, "ProfileImagePath", strProfileImagePath
				arrUserProfile = Split(strProfileImagePath, "\", -1, 1)
				Call Report(vbTab & vbTab & "> User: " & arrUserProfile(UBound(arrUserProfile)))
				If IsBound(arrUserProfileKey) Then
					' Profile is already loaded.
					'Call Report(vbTab & vbTab & "> Changing file from CURRENT_USER to " & arrUserProfile(UBound(arrUserProfile)))
					Call Report(vbTab & vbTab & "> Changing file from CURRENT_USER to " & USERS_REG & "\" & objProfile)
					ChangeRegFile CURRENT_USER_REG, USERS_REG & "\" & objProfile
					Call Report(vbTab & vbTab & vbTab & "# Adding registry settings from file: " & strOutputRegFile)
					intReturn = oWshShell.Run("Regedit /s """ & strOutputRegFile & """", 0, True)
				Else
					' Profile is not yet loaded.
					Call Report(vbTab & vbTab & "> Changing file from CURRENT_USER to " & USER_PROFILE_REG)
					ChangeRegFile CURRENT_USER_REG, USER_PROFILE_REG
					If Not IsNull(strProfileImagePath) Then
						Call Report(vbTab & vbTab & "> Profilepath: " & strProfileImagePath)
						If (UCase(strProfileImagePath) <> UCase(oWshPrsEnv("UserProfile"))) And oFSO.FileExists(strProfileImagePath & "\ntuser.dat") Then
							Call Report(vbTab & vbTab & vbTab & "> Loading Registry Hive into " & USER_PROFILE_REG)
							If LoadHive(strProfileImagePath) = 0 Then
								Call Report(vbTab & vbTab & vbTab & "> Adding registry settings from file " & strOutputRegFile)
								intReturn = oWshShell.Run("Regedit /s """ & strOutputRegFile & """", 0, True)
								Call Report(vbTab & vbTab & vbTab & "> Unloading User Registry Hive from " & USER_PROFILE_REG)
								UnloadHive
							Else
								Call Report(vbTab & vbTab & vbTab & "> Could not Load Hive (another hive may already be loaded)! Skipping...")
							End If
						ElseIf (UCase(strProfileImagePath) = UCase(oWshPrsEnv("UserProfile"))) Then
							Call Report(vbTab & vbTab & vbTab & "# Adding registry settings from file: " & strInputRegFile & " for the Current User " & oWshPrsEnv("UserProfile"))
							intReturn = oWshShell.Run("Regedit /s """ & strInputRegFile & """", 0, True)
						Else
							Call Report(vbTab & "# ntuser.dat not found for the curent profile " & oWshPrsEnv("UserProfile"))
						End If
					Else
						Call Report(vbTab & vbTab & vbTab & "> No Profilepath was defined. Skipping.")
					End If
				End If
			Else
				Call Report(vbTab & vbTab & "> Skipping Profile: " & objProfile)
			End If
		Next
	Else
		Call Report(vbTab & "> Could not find any profiles at HKLM\" & PROFILE_REG_BASE)
	End If
	
	Call Report(vbTab & "# Deleting temporary output file: " & strOutputRegFile)
	If oFSO.FileExists(strOutputRegFile) Then oFSO.DeleteFile strOutputRegFile, 1
	Call ExitScript("Function exit code: " & ApplyRegistryToProfiles)
	Set oFSO = Nothing
	Set oReg = Nothing
	Set oWshShell = Nothing
	Set oWshPrsEnv = Nothing
End Function

Private Function LoadHive(strProfilesPath)
	LoadHive = oWshShell.Run("Reg Load " & USER_PROFILE_REG & " """ & strProfilesPath & "\ntuser.dat""", 0, True)
End Function


Private Sub UnloadHive
	Dim intReturn
	intReturn = oWshShell.Run("Reg Unload " & USER_PROFILE_REG, 0, True)
End Sub

Private Sub ChangeRegfile(inChangeFrom, inChangeTo)
	Dim oInputReg, oOutputReg
	Dim strLine, strTempLine

	Set oInputReg = oFSO.OpenTextFile(strInputRegFile, CSTR_ForReading, False, -2)
	Set oOutputReg = oFSO.OpenTextFile(strOutputRegFile, CSTR_ForWriting, True)
	Do While Not oInputReg.AtEndOfStream
		strLine = oInputReg.ReadLine
		If Left(Trim(strLine), 1) = "[" And Right(Trim(strLine), 1) = "]" Then
			strLine = ChangeToDefaultUser(strLine, inChangeFrom, inChangeTo)
		End If
		oOutputReg.WriteLine strLine
	Loop
	oOutputReg.Close
	oInputReg.Close
	Set oOutputReg = Nothing
	Set oInputReg = Nothing
End Sub

Private Function ChangeToDefaultUser(inLine, inChangeFrom, inChangeTo)
	Dim strTemp
	Dim arrTemp

	strTemp = Mid(Trim(inLine), 2, Len(Trim(inLine)) -2)
	arrTemp = Split(strTemp, "\")
	If UBound(arrTemp) > 0 Then
		arrTemp(0) = Replace(arrTemp(0), inChangeFrom, inChangeTo)
	End If
	strTemp = "[" & Join(arrTemp, "\") & "]"
	ChangeToDefaultUser = strTemp
End Function

Private Function IsBound(inArray)
	On Error Resume Next

	If UBound(inArray) >= 0 Then
		IsBound = True
	Else
		IsBound = False
	End If
	If Err.Number <> 0 Then
		Err.Clear
		IsBound = False
	End If
End Function

Sub Report(ByVal msg)
	Const msiMessageTypeInfo = &H04000000
	Dim record
	
	If sFunctionString = "" Then sFunctionString = "CustomActionScriptLog"

    Set record = Installer.CreateRecord(1)
    record.stringdata(0) = sFunctionString & ": [1]"
    record.stringdata(1) = msg
    record.formattext
    message msiMessageTypeInfo, record
    Set record = Nothing
End Sub

Function ExitScript(msg)
	Call Report(msg)
	Call Report("Script end.")
	Call Report("####################################################################################################################")
End Function
