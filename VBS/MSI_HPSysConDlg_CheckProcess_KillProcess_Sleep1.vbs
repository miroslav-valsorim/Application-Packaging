'=================================================================================================================================
' Script: MSI_HPSysConDlg
' Version: 1.0
' Description: Checks for a pre-set process(es) and displays a message to the user if those processes need to be closed
'
' Usage: The script should be added as vbs custom action to MSI/MST
'	The process(es) are set in arrProcessesToCheck
'	Title and text of the dialog are set accordingly in nDialogReturn and sText
' 	VBS CA properties: "Stored in binary table", Target:HPSysConDialog, Synchronous (Check exit code), Immediate Execution, Always execute, <First Action>,
' 	Condition: NOT Installed OR REMOVE~="All" OR NOT NVD_VERIFY_MODE (so Radia doesn't execute it on Verify connect)
'
'=================================================================================================================================

Option Explicit

On Error Resume Next

Dim oWMI		' WMI object
Dim oDlg		' HPSysConDlgSrv.exe object
Dim oFS			' FileSystemObject

Dim nDlgReturn		' Return code from dialog
Dim sText			' Text to display in the dialog
Dim bLogFileOpen	' Indicates that LOG file is already open
Dim sFunctionString

Dim oProcess
Dim colProcess
Dim arrProcessesToCheck, sProcessToCheck
Dim i
Dim blnFound

' Init variables
' ========================================
bLogFileOpen = False
sFunctionString = ""
arrProcessesToCheck = Array("ami.exe")

' Create objects
' ========================================
Err.Clear
Set oWMI 	= GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set oFS 	= CreateObject("Scripting.FileSystemObject")
'If Err.Number <> 0 Then Call ExitScript("Error creating objects!", 3)

Function HPSysConDialog()
	sFunctionString = "HPSysConDialog"
	Call WriteToLog("####################################################################################################################")
	Call WriteToLog("Script start ...")

	' Create HpSysConDialog object
	' ========================================
	Call WriteToLog(vbTab & "Creating HpSysConDlg.HpDialog object...")
	Err.Clear
	Set oDlg = CreateObject("HpSysConDlg.HpDialog")
	If Err.Number <> 0 Then
		HpSysConDialog = ExitScript(vbTab & "Can't create HpSysConDlg.HpDialog object!", 3)
		Exit Function
	End If

	'Check for running processes before displaying the dialog
	' =======================================================================================================
	Call WriteToLog(vbTab & "Check for running processes before displaying the dialog.")

	i=0
	blnFound = False

	CheckProcesses()

	' Processes that might prevent Install are found. The initial dialog will be displayed
	If blnFound = True Then
		' Notify user to close applications before uninstallation
		' ========================================
		Call WriteToLog(vbTab & "Notify user that AMI Client application should be closed.")

		sText =  "Bitte, schließen Sie AMI Client Anwendung um die Installation oder Upgrade starten zu können." & vbCr _
					& vbCr & vbCr _
					& "Please close AMI Client application so that Uninstallation or Upgrade of the software can be started." & vbCr & vbCr _
					& "You have %d seconds to respond"
				
		nDlgReturn = 0
	
		nDlgReturn = oDlg.HpPopUp(sText, 300, "AMI Client Upgrade\Uninstall", 64)
		If Err.Number <> 0 Then
			HpSysConDialog = ExitScript(vbTab & "HpPopUp method returned error: " & Err.Number, 3)
			Exit Function
		End If
		Select Case nDlgReturn
			Case 1: 	Call WriteToLog(vbTab & "User confirmed!")
			Case "-1": 	Call WriteToLog(vbTab & "User didn't respond!")
			Case "-5": 	Call WriteToLog(vbTab & "Return code from HpPopup is -5! No user is logged on, so Install/Uninstall will continue!")
			Case Else: 	HpSysConDialog = ExitScript(vbTab & "HpPopUp returned error: " & nDlgReturn & " (" & HpSysConDlgSevError(nDlgReturn) & ")!", 3)
						Exit Function
		End Select
	Else
		sText=vbTab & "AMI Client process not found. Proceed with UnIstall\Upgrade!"
		HpSysConDialog = ExitScript(sText, 1)
		Exit Function
	End If

	sleep1 5

	' Check for running processes and ask user to close them
	' ===============================================================
	Call WriteToLog(vbTab & "Check for running processes and notify user which ones should be closed.")
	nDlgReturn = 1
	i=0
	blnFound = False

	While 1
		blnFound = False
		CheckProcesses()
		If i >= 2 And blnFound = True Then
			HpSysConDialog = ExitScript(vbTab & "AMI Client process found and killed. Proceed with UnIstall\Upgrade!", 1)
			Exit Function
		End If

		If blnFound = True Then
			Call WriteToLog(vbTab & "Notify user that AMI Client application should be closed.")
			If i >= 1 Then sText = stext & vbCr & vbCr & "WARNING!!! All relevant programs will be automatically closed after you press Ok!" & vbCr & vbCr _
							& "WARNUNG: Alle entsprechenden Programme werden nach der Betätigung der Ok-Taste automatisch beendet!"
			nDlgReturn = oDlg.HpPopUp(sText, 300, "AMI Client Upgrade Uninstall\Upgrade", 64)
			If Err.Number <> 0 Then
				HpSysConDialog = ExitScript(vbTab & "HpPopUp method returned error: " & Err.Number, 3)
				Exit Function
			End If
			Select Case nDlgReturn
				Case 1: 	Call WriteToLog(vbTab & "User confirmed!")
				Case "-1": 	If i < 1 Then
								Call WriteToLog(vbTab & "User didn't respond!")
							Else
								Call WriteToLog(vbTab & "User didn't respond. Killing all relevant applications!")
							End If
				Case "-5": 	Call WriteToLog(vbTab & "Return code from HpPopup is -5! No user is logged on, so Install/Uninstall will continue!")
				Case Else: 	HpSysConDialog = ExitScript(vbTab & "HpPopUp returned error: " & nDlgReturn & " (" & HpSysConDlgSevError(nDlgReturn) & ")!", 3)
							Exit Function
			End Select
			i = i+1
		Else
			HpSysConDialog = ExitScript(vbTab & "AMI Client process not found. Proceed with UnIstall\Upgrade!", 1)
			Exit Function
		End If
		
		sleep1 5
	Wend

	HpSysConDialog = ExitScript (vbTab & "Ending script.", 3)
End Function

' ===============================================================
' Return error description from HpSysConDlgSrv return code
' ===============================================================
Function HpSysConDlgSevError(nDialogReturn)
	Select Case nDialogReturn
		Case "-2": HpSysConDlgSevError = "Error occurred within exe"
		Case "-3": HpSysConDlgSevError = "Dll could not open a connection to the exe"
		Case "-4": HpSysConDlgSevError = "Invalid return from exe"
		Case "-5": HpSysConDlgSevError = "No users are logged in"
		Case "-6": HpSysConDlgSevError = "User logged in, but no exe running"
		Case Else: HpSysConDlgSevError = "Unknown error"
	End Select
End Function

' ===============================================================
' Write sExitText to log and quit, returning nExitCode
' ===============================================================
Function ExitScript(sExitText, nExitCode)
	Call WriteToLog(sExitText & vbTab & "ExitCode: " & nExitCode)
	Call WriteToLog("Script end.")
	Call WriteToLog("####################################################################################################################")
	ExitScript = nExitCode
	'WScript.Quit(nExitCode)
End Function

' ========================================
' Write to LOG file
' ========================================
Sub WriteToLog(sTextToWrite)
	'Write to MSI log
	' ========================================
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

' ========================================
' Check Processes
' ========================================
Sub CheckProcesses()
	For Each sProcessToCheck In arrProcessesToCheck
		Set colProcess = oWMI.ExecQuery("Select * from Win32_Process Where Name = """ & sProcessToCheck & """")
		If colProcess.Count > 0 Then
			Call WriteToLog(vbTab & vbTab & sProcessToCheck & " found!")
			If i >= 2 Then 
				KillProcesses(sProcessToCheck)
			End If
			blnFound = True
		End If 
	Next
End Sub
	
' ========================================
' Kill Process
' ========================================
Sub KillProcesses(ProcessToKill)
	On Error Resume Next
	Dim colProcess, oProcess
	Set colProcess = oWMI.ExecQuery("Select * from Win32_Process Where Name = """ & ProcessToKill & """")
	If colProcess.Count > 0 Then
		For Each oProcess in colProcess
			Call WriteToLog(vbTab & vbTab & ProcessToKill & " found! Killing!")
			oProcess.Terminate()
			sleep1 1
		Next
	End If
End Sub

' ========================================
' Wscript.Sleep alternative
' ========================================
Sub sleep1(strSeconds)
       Dim dteWait : dteWait = DateAdd("s", strSeconds, Now())
       Do Until (Now() > dteWait)
       Loop
End Sub