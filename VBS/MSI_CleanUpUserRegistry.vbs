'=================================================================================================================================
' Script: MSI_CleanUpUserRegistry.vbs
' Version: 1.0
' Description: The script is based on DefaultUser.vbs, but reworked to be executed as CA from MSI and to delete only keys set in arrKeysToDelete
'
' USAGE: The script must be executed from MSI VBS CA. Writes to log and returns appropriate return codes, so the CA must be 
' 		VBS-Stored in the Binary Table. Script Function: UserCleanUp. Condition: as required.
'		Warning!!! If Session.Property is used, like in this example, most of the properties are not visible in Deferred Execution, 
'		so Immediate Execution must be used for the CA
'
'=================================================================================================================================

Option Explicit

CONST	HKEY_CURRENT_USER 	= &H80000001
CONST	HKEY_LOCAL_MACHINE 	= &H80000002
CONST	HKEY_USERS	 	= &H80000003

CONST	USER_PROFILE_REG	= "HKEY_LOCAL_MACHINE\Temp"
CONST	PROFILE_REG_BASE	= "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

Private oFSO, oWshShell, oWshPrsEnv, oReg
Private intReturn

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWshShell = CreateObject("WScript.Shell")
Set oWshPrsEnv = oWshShell.Environment("PROCESS")
Set oReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!//" & oWshPrsEnv("ComputerName") & "/root/default:StdRegProv")

Dim arrKeysToDelete, sFunctionString

arrKeysToDelete = Array("\Software\Wow6432Node\Microsoft\Active Setup\Installed Components\{E8DCB7E1-31A1-4956-BFE5-46FA177AC853}", _
						"\\Software\Adobe\CommonFiles" _
						)

' arrKeysToDelete = Array("\Software\Wow6432Node\Microsoft\Active Setup\Installed Components\" & Session.Property("ProductCode"))

sFunctionString = ""

Function UserCleanUp()
	On Error Resume Next

	Dim arrProfiles
	Dim objProfile
	Dim strProfileImagePath
	Dim arrUserProfileKey
	Dim arrUserProfile
	Dim arrKeysToDelete
	
	sFunctionString = "UserCleanUp"
	
	UserCleanUp = 1
	Call Report("####################################################################################################################")
	Call Report("# Applying registry to Profiles List.")

	' All profiles are enumerated
	Call Report(vbTab & "> Enumerating Profiles")
	oReg.EnumKey HKEY_LOCAL_MACHINE, PROFILE_REG_BASE, arrProfiles
	If IsBound(arrProfiles) Then
		For Each objProfile in arrProfiles
			If Err.Number <> 0 Then
				UserCleanUp = 3
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
					DeleteEntry HKEY_USERS, objProfile, strProfileImagePath
				Else
					' Profile is not yet loaded.
					If Not IsNull(strProfileImagePath) Then
						Call Report(vbTab & vbTab & "> Profilepath: " & strProfileImagePath)
						If (UCase(strProfileImagePath) <> UCase(oWshPrsEnv("UserProfile"))) And oFSO.FileExists(strProfileImagePath & "\ntuser.dat") Then
							Call Report(vbTab & vbTab & vbTab & "> Loading Registry Hive into " & USER_PROFILE_REG)
							If LoadHive(strProfileImagePath) = 0 Then
								DeleteEntry HKEY_LOCAL_MACHINE, "Temp", strProfileImagePath
								Call Report(vbTab & vbTab & vbTab & "> Unloading User Registry Hive from " & USER_PROFILE_REG)
								UnloadHive
							Else
								Call Report(vbTab & vbTab & vbTab & "> Could not Load Hive (another hive may already be loaded)! Skipping...")
							End If
						ElseIf (UCase(strProfileImagePath) = UCase(oWshPrsEnv("UserProfile"))) Then
							DeleteEntry HKEY_CURRENT_USER, "", strProfileImagePath
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
	
	Call ExitScript("Function exit code: " & UserCleanUp)
End Function


Private Function LoadHive(strProfilesPath)
	LoadHive = oWshShell.Run("Reg Load " & USER_PROFILE_REG & " """ & strProfilesPath & "\ntuser.dat""", 0, True)
End Function


Private Sub UnloadHive
	Dim intReturn
	intReturn = oWshShell.Run("Reg Unload " & USER_PROFILE_REG, 0, True)
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
	
	For Each strKey In arrKeysToDelete
		strRegPath = strProfileToSearch & strKey
		Call Report(vbTab & vbTab & vbTab & "> Checking for Path: """ & strRegPath & """ in User: " & strProfilePath)
		DeleteSubkeys iHiveToSearch, strRegPath
	Next
End Function


Sub DeleteSubkeys(HKEY_HIVE, strKeyPath) 
	Dim arrSubKeys, strSubkey
    oReg.EnumKey HKEY_HIVE, strKeyPath, arrSubkeys 
    If IsArray(arrSubkeys) Then 
        For Each strSubkey In arrSubkeys 
            DeleteSubkeys HKEY_HIVE, strKeyPath & "\" & strSubkey
        Next
    End If 
    intReturn = oReg.DeleteKey(HKEY_HIVE, strKeyPath)
	If (intReturn = 0) And (Err.Number = 0) Then Call Report(vbTab & vbTab & vbTab & "> " & strKeyPath & " deleted")
End Sub


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
	Call Report("####################################################################################################################")
End Function
