'************************************************************************************************************************
'*		Script: DefaultUser.vbs											 *
'*		Version: 1.0.3												*
'*		Author(s): HP Flexdeploy Development Team								*
'*															*
'*		Usage: 	DefaultUser.vbs script requires a parameter pointing to a registry file (including path).	*
'*			This file will then be parsed where all the HKEY_CURRENT_USER settings will be changed		*
'*			to point to a DEFAULT USER registry key (which can be changed using the USER_PROFILE_REG	*
'*			constant).											*
'*			Returns error 10 when no parameter (registry file) is defined.					*
'*			Returns error 20 when the registry file does not exist.						*
'*															*
'*															*
'*		V1.0.1:	Changed enumeration of all users in the WMI, to an enumeration of the ProfilesList.		*
'*		v1.0.2: Added the Loadhive for the Default User. This user was not handled.				*
'*		v1.0.3:	Added enumeration of HKEY_USERS to keep track of hives that are already loaded 			*
'*			(user is no longer required to log off).							*
'*		v1.0.4: Changed behaviour to look up the Default User's path, as this has changed with Vista.		*
'*															*
'*		Copyright (C) 2002-2003 Hewlett Packard									*
'*															*
'************************************************************************************************************************

Option Explicit

SetLocale(1033)
CONST   CSTR_VERSION            = "1.0.3"
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

Private oFSO, oWshShell, oWshSysEnv, oWshPrsEnv, oWMI, oReg
Private GSTR_LOGFILE
Private strInputRegFile, strOutputRegFile, strFileName
Private arrSplitFilePath
Private oItem
Private intReturn

Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set oWshShell = WScript.CreateObject("WScript.Shell")
Set oWshPrsEnv = oWshShell.Environment("PROCESS")
Set oReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!//" & oWshPrsEnv("ComputerName") & "/root/default:StdRegProv")

GSTR_LOGFILE = oWshPrsEnv("FlxLogFilesPath")
If GSTR_LOGFILE = vbNullString Then
	If oWshPrsEnv("TEMP") <> vbNullString Then
		GSTR_LOGFILE = oWshPrsEnv("TEMP") & "\DefaultUser.Log"
	Else
		GSTR_LOGFILE = oWshShell.CurrentDirectory & "\DefaultUser.Log"
	End If
Else
	GSTR_LOGFILE = GSTR_LOGFILE & "\DefaultUser.Log"
End If

Call Report("# Starting " & Wscript.ScriptName & " version " & CSTR_VERSION)

Call Report("# Checking for parameters.")
If WScript.Arguments.Count <> 1 Then
	Call ExitOnError(10, "> No parameters found.")
Else
	strInputRegFile = WScript.Arguments(0)
	Call Report(vbTab & "> Parameter found. Using " & strInputRegFile & " as input file.")
End If
Call Report("# Checking file existence.")
If oFSO.FileExists(strInputRegFile) Then
	arrSplitFilePath = Split(strInputRegFile, "\")
	If UBound(arrSplitFilePath) >= 1 Then
		strFileName = DEFAULT_USER_FILE & arrSplitFilePath(UBound(arrSplitFilePath))
		If Right(strFileName, 1) = """" Then
			strFileName = Mid(strFileName, 1, Len(strFileName) -1)
		End If
	Else
		strFileName = DEFAULT_USER_FILE & strInputRegFile
	End If
	If oWshPrsEnv("TEMP") <> vbNullString Then
		strOutputRegFile = oWshPrsEnv("TEMP") & "\"
	Else
		arrSplitFilePath = Split(WScript.ScriptFullName, "\")
		If UBound(arrSplitFilePath) >= 1 Then
			arrSplitFilePath(UBound(arrSplitFilePath)) = vbNullString
			strOutputRegFile = Join(arrSplitFilePath, "\")
		End If
	End If
	strOutputRegFile = strOutputRegFile & strFileName
	Call Report(vbTab & "> Inputfile found. Using " & strOutputRegFile & " as output file.")
Else
	Call ExitOnError(20, "Inputfile not found.")
End If
' Call Report("# Changing file from CURRENT_USER to DEFAULT_USER")
' ChangeRegFile CURRENT_USER_REG, USER_PROFILE_REG

Call Report("# Applying registry to Profiles List.")
ApplyRegistryToProfiles

Call Report("# Ending " & Wscript.ScriptName & " version " & CSTR_VERSION)
Call Report("")

Set oFSO = Nothing
Set oReg = Nothing
Set oWshShell = Nothing
Set oWshSysEnv = Nothing
Set oWshPrsEnv = Nothing


Private Function LoadHive(strProfilesPath)

	LoadHive = oWshShell.Run("Reg Load " & USER_PROFILE_REG & " """ & strProfilesPath & "\ntuser.dat""", 0, True)

End Function


Private Sub UnloadHive

	Dim intReturn

	intReturn = oWshShell.Run("Reg Unload " & USER_PROFILE_REG, 0, True)

End Sub


Private Sub ApplyRegistryToProfiles

	On Error Resume Next

	Dim arrProfiles
	Dim objProfile
	Dim strProfileImagePath
	Dim strDefaultUserProfile, strDefaultUserPath, strWindowsVersion
	Dim arrUserProfileKey
	Dim objUserProfileKey
	Dim arrUserProfile

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
				Call ExitOnError(Err.Number, vbTab & "> " & Err.Description)
				Exit For
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
					Call Report(vbTab & vbTab & "> Changing file from CURRENT_USER to " & arrUserProfile(UBound(arrUserProfile)))
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


Private Function Report(text)

	Dim oFW
	Dim sComputerName

	sComputerName = oWshPrsEnv("Computername")

	Wscript.Echo now & vbTab & sComputerName & vbTab & text

	Set oFW = oFSO.OpenTextFile(GSTR_LOGFILE, CSTR_ForAppending, True)
	oFW.WriteLine now & vbTab & sComputerName & vbTab & text
	oFW.Close
	Set oFW = Nothing

End Function


Private Sub ExitOnError(iCode,sDescription)

	Call Report(vbTab & "An error occured: " & sDescription)
	Call Report("# Exiting " & Wscript.ScriptName & " prematurely with exitcode " & iCode)
	Wscript.Quit(iCode)

End Sub
