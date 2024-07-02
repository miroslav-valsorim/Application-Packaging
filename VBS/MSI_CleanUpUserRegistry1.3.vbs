'=================================================================================================================================
' Script: MSI_CleanUpUserRegistry.vbs
' Version: 1.3
' Description: The script is based on DefaultUser.vbs, but reworked to be executed as CA from MSI and to delete only keys set in arrKeysToDelete and/or values set in arrValuesToDelete
'
' USAGE: The script must be executed from MSI VBS CA. Writes to log and returns appropriate return codes, so the CA must be 
' 		VBS-Stored in the Binary Table. Script Function: UserCleanUp. Condition: as required.
'		Warning!!! If Session.Property is used, like in this example, most of the properties are not visible in Deferred Execution, 
'		so Immediate Execution must be used for the CA
'
' CHANGELOG: 02.08.2017 - Added ability to delete keys from "HKEY_CURRENT_USER\Software\Classes" by loading Classes hives from %LocalAppData%\Microsoft\Windows\UsrClass.dat
'			 10.07.2019 - Added arrValuesToDelete 2D Array with first column registry key and second column registry value to delete
'			 12.12.2020 - Fixed issue with DeleteEntry function
'=================================================================================================================================

Option Explicit

CONST	HKEY_CURRENT_USER 	= &H80000001
CONST	HKEY_LOCAL_MACHINE 	= &H80000002
CONST	HKEY_USERS	 	= &H80000003

CONST	USER_PROFILE_REG	= "HKEY_LOCAL_MACHINE\Temp" ' USER_PROFILE_REG and USER_PROFILE_REG & "_Classes" are the locations where user registry and classes user registry are loaded
														' The unload function assumes the hives are loaded to "HKEY_LOCAL_MACHINE". If any other root key is used, the "Hive Loaded Check" will fail
CONST	PROFILE_REG_BASE	= "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
CONST	WINDOWSNT_VERSION_REG	= "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

Private oFS, oShell, oWshPrsEnv, oReg
Private intReturn

Set oFS = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
Set oWshPrsEnv = oShell.Environment("PROCESS")
Set oReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!//" & oWshPrsEnv("ComputerName") & "/root/default:StdRegProv")

Dim arrKeysToDelete, arrValuesToDelete
Dim sFunctionString : sFunctionString = ""
Dim bClassesFound : bClassesFound = 0
Dim intRegLoaded : intRegLoaded = 1 ' This is the return code from Reg Load command. 0 - registry successfully loaded, 1 - failed to load
Dim intClassesLoaded : intClassesLoaded = 1 ' This is the return code from Reg Load command. 0 - registry successfully loaded, 1 - failed to load

' Example with Session Properties
'arrKeysToDelete = Array("\Software\" & Session.Property("Manufacturer") & "\" & Session.Property("DevTrackID") & "_" & Session.Property("ProductName"), _
'						"\Software\Wow6432Node\Microsoft\Active Setup\Installed Components\" & Session.Property("ProductCode"))

' arrKeysToDelete = Array("\Software\Microsoft\Office\14.0\Outlook\Addins\EnterpriseVault.DesktopUI", _
						' "\Software\Microsoft\Office\15.0\Outlook\Addins\EnterpriseVault.DesktopUI", _
						' "\Software\Microsoft\Office\16.0\Outlook\Addins\EnterpriseVault.DesktopUI", _
						' "\Software\KVS\Enterprise Vault" _
						' )

' arrValuesToDelete = Array(_
							' Array("\Software\Microsoft\Office\14.0\Outlook\AddInLoadTimes", "EnterpriseVault.DesktopUI"), _
							' Array("\Software\Microsoft\Office\15.0\Outlook\AddInLoadTimes", "EnterpriseVault.DesktopUI"), _
							' Array("\Software\Microsoft\Office\16.0\Outlook\AddInLoadTimes", "EnterpriseVault.DesktopUI") _
							' )


If Ubound(Filter(arrKeysToDelete, "\SOFTWARE\Classes\", True, vbTextCompare)) > -1 Then bClassesFound = 1

Function UserCleanUp()
	On Error Resume Next

	Dim arrProfiles
	Dim objProfile
	Dim strProfileImagePath
	Dim strDefaultUserProfile, strDefaultUserPath, strWindowsVersion
	Dim arrUserProfileKey
	Dim arrUserProfile
	
	sFunctionString = "UserCleanUp"
	
	UserCleanUp = 1
	Call WriteToLog("####################################################################################################################")
	Call WriteToLog("# Applying registry to Profiles List.")

	' First we load the hive from the Default User and change the corresponding registry settings.
	Call WriteToLog("> Checking Default User")
	oReg.GetStringValue HKEY_LOCAL_MACHINE, WINDOWSNT_VERSION_REG, "CurrentVersion", strWindowsVersion
	If Left(strWindowsVersion, 1) > 5 Then
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, "Default", strProfileImagePath
	Else
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, "ProfilesDirectory", strDefaultUserPath
		oReg.GetStringValue HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, "DefaultUserProfile", strDefaultUserProfile
		strProfileImagePath = strDefaultUserPath & "\" & strDefaultUserProfile
	End If
	If oFS.FileExists(strProfileImagePath & "\ntuser.dat") Then
		LoadHive strProfileImagePath
		DeleteEntry HKEY_LOCAL_MACHINE, "Temp", strProfileImagePath
		UnloadHive
	Else
		Call WriteToLog(vbTab & "# ntuser.dat not found for the Default user: " & strProfileImagePath & "\ntuser.dat")
	End If

	' Next, all profiles are enumerated
	Call WriteToLog("> Enumerating Profiles")
	oReg.EnumKey HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, arrProfiles
	'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList - list of keys with user profiles
	If IsBound(arrProfiles) Then
		For Each objProfile in arrProfiles ' objProfile - User SID, e.g. S-1-5-21-2925477919-3492066975-685244884-500 for Administrator
			If Err.Number <> 0 Then
				UserCleanUp = 3
				Exit Function
			End If
			' UserProfiles start with S-1-5-21-... All other profiles, like Network and System, are ignored.
			If Instr(1, objProfile, "S-1-5-21-", vbTextCompare) Then
				Call WriteToLog("# Checking Profile: " & objProfile)
				oReg.EnumKey HKEY_USERS, objProfile, arrUserProfileKey
				' Check whether the profile is already loaded.
				oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, PROFILE_REG_BASE & "\" & objProfile, "ProfileImagePath", strProfileImagePath ' e.g. strProfileImagePath=C:\Users\Administrator
				arrUserProfile = Split(strProfileImagePath, "\", -1, 1)
				Call WriteToLog("> User: " & arrUserProfile(UBound(arrUserProfile)))
				If IsBound(arrUserProfileKey) Then
					' Profile is already loaded - SID key is found in HKU
					DeleteEntry HKEY_USERS, objProfile, strProfileImagePath
				Else
					' Profile is not yet loaded.
					If Not IsNull(strProfileImagePath) Then
						Call WriteToLog("> Profilepath: " & strProfileImagePath)
						If (UCase(strProfileImagePath) <> UCase(oWshPrsEnv("UserProfile"))) And oFS.FileExists(strProfileImagePath & "\ntuser.dat") Then
							LoadHive strProfileImagePath
							DeleteEntry HKEY_LOCAL_MACHINE, "Temp", strProfileImagePath
							UnloadHive
						ElseIf (UCase(strProfileImagePath) = UCase(oWshPrsEnv("UserProfile"))) Then
							' The profile wasn't found as loaded, but the profile path is the CU path, e.g. C:\Users\Administrator
							DeleteEntry HKEY_CURRENT_USER, "", strProfileImagePath
						Else
							Call WriteToLog(vbTab & "# ntuser.dat not found for the current profile " & strProfileImagePath)
						End If
					Else
						Call WriteToLog(vbTab & "> No Profilepath was defined. Skipping.")
					End If
				End If
			Else
				Call WriteToLog("> Skipping Profile: " & objProfile)
			End If
		Next
	Else
		Call WriteToLog("> Could not find any profiles at HKLM\" & PROFILE_REG_BASE)
	End If
	
	Call ExitScript("Function exit code: " & UserCleanUp)
End Function


Private Sub LoadHive(strProfilesPath)
	intRegLoaded = 1
	intClassesLoaded = 1
	
	Call WriteToLog(vbTab & "> Loading Registry Hive into " & USER_PROFILE_REG)
	intRegLoaded = oShell.Run("Reg Load " & USER_PROFILE_REG & " """ & strProfilesPath & "\ntuser.dat""", 0, True)
	If bClassesFound Then
		If oFS.FileExists(strProfilesPath & "\AppData\Local\Microsoft\Windows\UsrClass.dat") Then
			Call WriteToLog(vbTab & "> Loading Classes Registry Hive into " & USER_PROFILE_REG & "_Classes")
			intClassesLoaded = oShell.Run("Reg Load " & USER_PROFILE_REG  & "_Classes" & " """ & strProfilesPath & "\AppData\Local\Microsoft\Windows\UsrClass.dat""", 0, True)
			If intClassesLoaded <> 0 Then intRegLoaded = intClassesLoaded
		Else
			Call WriteToLog(vbTab & strProfilesPath & "\AppData\Local\Microsoft\Windows\UsrClass.dat doesn't exist. No Classes registry keys will be deleted")
		End If
	End If
	
	If intRegLoaded <> 0 Then
		Call WriteToLog(vbTab & "> Could not Load Hive!")
		Call ExitScript("Fatal error", 3)
	End If
End Sub


Private Sub UnloadHive
	Dim intRegUnload: intRegUnload = 1
	Dim intClassesUnload: intRegUnload = 1
	Dim arrLoadedHiveCheck

	oReg.EnumKey HKEY_LOCAL_MACHINE, Replace(USER_PROFILE_REG, "HKEY_LOCAL_MACHINE\", "", 1, 1, vbTextCompare), arrLoadedHiveCheck
	If IsBound(arrLoadedHiveCheck) Then
		Call WriteToLog(vbTab & "> Unloading User Registry Hive from " & USER_PROFILE_REG)
		intRegUnload = oShell.Run("Reg Unload " & USER_PROFILE_REG, 0, True)
	Else
		intRegUnload = 0
	End If
	
	oReg.EnumKey HKEY_LOCAL_MACHINE, Replace(USER_PROFILE_REG, "HKEY_LOCAL_MACHINE\", "", 1, 1, vbTextCompare) & "_Classes", arrLoadedHiveCheck
	If IsBound(arrLoadedHiveCheck) Then
		Call WriteToLog(vbTab & "> Unloading Classes Registry Hive from " & USER_PROFILE_REG & "_Classes")
		intClassesUnload = oShell.Run("Reg Unload " & USER_PROFILE_REG & "_Classes", 0, True)
		If intClassesUnload <> 0 Then intRegUnload = intClassesUnload
	Else
		intClassesUnload = 0
	End If
	
	If intRegUnload <> 0 Then
		Call WriteToLog(vbTab & "> Could not Unload Hive!")
		Call ExitScript("Fatal error", 3)
	End If
End Sub


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


Function DeleteEntry(iHiveToSearch, strProfileToSearch, strProfilePath)
	Dim strRegPath
	Dim strKey

	If IsBound(arrKeysToDelete) Then
		For Each strKey In arrKeysToDelete
			strRegPath = strProfileToSearch & strKey

			If Instr(1, strRegPath, "\Software\Classes", vbTextCompare) Then strRegPath = Replace(strRegPath, "\SOFTWARE\Classes", "_Classes", 1, 1, vbTextCompare)
		
			Call WriteToLog(vbTab & "> Checking for Path: """ & strRegPath & """ in User: " & strProfilePath)
			If RegExists(iHiveToSearch, strRegPath, "") Then
				DeleteSubkeys iHiveToSearch, strRegPath
			Else
				Call WriteToLog(vbTab & "> """ & strRegPath & """ doesn't exist")
			End If
		Next
	End If

	If IsBound(arrValuesToDelete) Then
		For Each strKey In arrValuesToDelete
			strRegPath = strProfileToSearch & strKey(0)
			
			If Instr(1, strRegPath, "\Software\Classes", vbTextCompare) Then strRegPath = Replace(strRegPath, "\SOFTWARE\Classes", "_Classes", 1, 1, vbTextCompare)

			Call WriteToLog(vbTab & "> Checking for Path: """ & strRegPath & ":  " & strKey(1) & """ in User: " & strProfilePath)
			If RegExists(iHiveToSearch, strRegPath, strKey(1)) Then
				intReturn = oReg.DeleteValue(iHiveToSearch, strRegPath, strKey(1))
				If (intReturn = 0) And (Err.Number = 0) Then
					Call WriteToLog(vbTab & "> " & strRegPath & ":  " & strKey(1) & " deleted")
				Else
					Call WriteToLog(vbTab & "> Error deleting " & strRegPath & ":  " & strKey(1) & ". Return Code: " & intReturn)
				End If
			Else
				Call WriteToLog(vbTab & "> """ & strRegPath & ":  " & strKey(1) & """ doesn't exist")
			End If
		Next
	End If
End Function


' ===============================================================
' DeleteSubkeys recursively  deletes the provided key and all subkeys
' ===============================================================
Sub DeleteSubkeys(HKEY_HIVE, strKeyPath) 
	Dim arrSubKeys, strSubkey
	Dim iReturn
	Dim sHKEY_HIVE
	
	sHKEY_HIVE = TranslateHive(HKEY_HIVE)
    
	oReg.EnumKey HKEY_HIVE, strKeyPath, arrSubkeys 
    If IsArray(arrSubkeys) Then 
        For Each strSubkey In arrSubkeys 
            DeleteSubkeys HKEY_HIVE, strKeyPath & "\" & strSubkey
        Next
    End If
	
	If RegExists(HKEY_HIVE, strKeyPath, "") Then
		iReturn = oReg.DeleteKey(HKEY_HIVE, strKeyPath)
		If (iReturn = 0) And (Err.Number = 0) Then
			Call WriteToLog(vbTab & vbTab & sHKEY_HIVE & "\" & strKeyPath & " deleted")
		Else
			Call WriteToLog(vbTab & vbTab & "Error deleting " & sHKEY_HIVE & "\" & strKeyPath & ". Return Code: " & iReturn)
		End If
	Else
		Call WriteToLog(vbTab & vbTab & sHKEY_HIVE & "\" & strKeyPath & " doesn't exist")
	End If
End Sub	'End DeleteSubkeys

' ===============================================================
' RegExists Returns True if registry key or value exists.
' Set strValue="" if you want to check key existence
' ===============================================================
Function RegExists(ByVal HKEY_HIVE, ByVal strKeyPath, strValue)
	On Error Resume Next
	
	Dim sHKEY_HIVE

	RegExists = True
	
	sHKEY_HIVE = TranslateHive(HKEY_HIVE)

	If Left(strKeyPath, 1) <> "\" Then strKeyPath = "\" & strKeyPath
	strKeyPath = sHKEY_HIVE & strKeyPath
	If Right(strKeyPath, 1) <> "\" Then strKeyPath = strKeyPath & "\"
	If strValue <> "" Then strKeyPath = strKeyPath & strValue
	Err.Clear
	oShell.RegRead(strKeyPath)
	If Err <> 0 Then RegExists = False
	Err.Clear
End Function 'RegExists


' ===============================================================
' TranslateHive
' ===============================================================
Function TranslateHive(HKEY_HIVE)
	If IsNumeric(HKEY_HIVE) Then
		Select Case HKEY_HIVE
			Case &H80000003: TranslateHive = "HKEY_USERS"
			Case &H80000002: TranslateHive = "HKEY_LOCAL_MACHINE"
			Case &H80000001: TranslateHive = "HKEY_CURRENT_USER"
		End Select
	Else
		TranslateHive = HKEY_HIVE
	End If
End Function 'TranslateHive


Sub WriteToLog(ByVal msg)
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
	Call WriteToLog(msg)
	Call WriteToLog("####################################################################################################################")
End Function