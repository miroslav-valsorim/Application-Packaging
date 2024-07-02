'==========================================================================
' Name: CheckStartStopService.vbs
' Version: 1.0
' Description: A function that can be used to check the state of a service, start or stop it
'
' Usage: Call CheckStartStopService(sMode, sServiceName) with the following parameters:
'	sMode: "Check"; "Start"; "Stop"
'	sServiceName: Name of the service. Service name (not Display Name) of the service must be used
'	The function returns:
'		If "Check" mode is used: state of the service
' 		If "Start" or "Stop" mode is used: returns 1 if successful , or 0 otherwise
'==========================================================================

Option Explicit

On Error Resume Next

msgbox CheckStartStopService("Check", "W32Time")

' ===============================================================
' CheckStartStopService
' Valid sMode: "Check"; "Start"; "Stop"
' If "Check" return State of the service
' If "Start" or "Stop" return 1 if success, or 0 otherwise
' ===============================================================
Function CheckStartStopService(sMode, sServiceName)
	Dim oWMIService, colListOfServices, oService
	
	CheckStartStopService = 0
	Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colListOfServices = oWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & sServiceName & "'")
	If colListOfServices.count > 0 Then
		For Each oService in colListOfServices
			Select Case sMode
				Case "Check":	CheckStartStopService = oService.State
								MsgBox oService.DisplayName & " state: " & CheckStartStopService
				Case "Start":	iReturn = oService.StartService()
								If iReturn = 0 Then
									MsgBox oService.DisplayName & " started successfully"
									CheckStartStopService = 1
								ElseIf iReturn = 10 Then 
									MsgBox oService.DisplayName & " is already running"
									CheckStartStopService = 1
								Else
									MsgBox "Failed to start " & oService.DisplayName & ". Return from StartService:" & iReturn
									CheckStartStopService = 0
								End If
				Case "Stop":	iReturn = oService.StopService()
								If iReturn = 0 Then
									MsgBox oService.DisplayName & " stopped successfully"
									CheckStartStopService = 1
									
								ElseIf iReturn = 5 Then
									MsgBox oService.DisplayName & " is already stopped"
									CheckStartStopService = 1
								Else
									MsgBox "Failed to stop " & oService.DisplayName & ". Return from StopService:" & iReturn
									CheckStartStopService = 0
								End If
				Case Else:		MsgBox "Invalid sMode passed"
								CheckStartStopService = 0
								Exit For
			End Select
		Next
	Else
		If IsObject(oService) Then
			MsgBox "Can't find " & oService.DisplayName
		Else
			MsgBox "Can't find " & sServiceName & " service"
		End If
	End If
End Function