'********************************************************************
'*
'*  Name:            CSI_GetSDDLFromObject.vbs
'*  Author:          Darwin Sanoy
'*  Updates:         http://csi-windows.com/community
'*  Bug Reports &
'   Enhancement Req: http://csi-windows.com/contactus
'*
'*  License:  If a company has sent someone to our training, they may use this  
'*            script on all computers in that company.  Companies who have not
'*            sent anyone to our training can use the above contact link to 
'*            request licensing information.
'*
'*  Built/Tested On: Windows 7
'*  Requires:        OS: Windows Vista, Windows 7, Server 2008
'*
'*  Main Function:
'*     Retrieve an SDDLText identifier from an existing file, folder, 
'*     registry key or service. 
'*
'*  Syntax:
'*    Interactive (including prompt for location and clipboard support): 
'*
'*        1) "wscript CSI_GetSDDLFromObject.vbs [Filepath | RegPath | service:Service Name]" (CTRL-C to copy SDDLText)
'*        2) command prompt "cscript CSI_GetSDDLFromObject.vbs [Filepath | RegPath]"
'*     
'*    The script takes either type of registry reference 
'*       (e.g. "HKLM" OR "HKEY_LOCAL_MACHINE")
'* 
'*    The script takes either type of service name 
'*       (e.g. "Adaptive Brightness" OR "SensrSvc")
'*
'*  Usage and Notes:
'*    
'*  Implementation Details: 
'* 
'*  Assumptions &
'*  Limitations:
'*
    Const SCRIPTVERSION = 1.2
'*
'*  Revision History:
'*     11/05/10 - 1.2 - updated to handle services (djs)
'*     11/04/10 - 1.1 - inital version (djs)
'*
'*******************************************************************

Const WMI_HKLM = &H80000002
Const WMI_HKCU = &H80000001
Set oShell = WScript.CreateObject("WScript.Shell")
Set oWMI = GetObject("winmgmts:\\.\root\default")
Set oWMICIMV2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
set oReg = oWMI.Get("StdRegProv")
set sdHelper = oWMICIMV2.Get("Win32_SecurityDescriptorHelper")

bDebugMessagesOn = True

UserMessage "'**********************************************" &vbcrlf _
& "'*  " &vbcrlf _
& "'*   Script: " & wsh.scriptname &vbcrlf _
& "'*  Version: " & SCRIPTVERSION &vbcrlf _
& "'*  Updates: http://CSI-Windows.com/toolkit" &vbcrlf _
& "'*  " &vbcrlf _
& "'**********************************************" &vbcrlf

REM If NOT cint(oShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")) > 5 Then
  REM DebugMessage "XP does not have the required WMI classes for this script, exiting..." 
REM End If 

If NOT isAdmin() Then
  wscript.echo "You must be an administrator to run this script - run from " _
  & "an elevated command line with wscript for interactive and cscript for " _
  & "commandline mode, exiting..."
  wsh.quit
End If

If (Instr(1,Wsh.FullName,"wscript.exe",1) > 0) Then bInteractive = True

if wsh.arguments.count = 0 Then
  If bInteractive Then
    Resource = inputbox("Extract SDDLText From Resource" _ 
    & vbcrlf & vbcrlf & "Please type in the full path of the file, folder or " _
    & "registry Key (without value name) or 'Service:' followed by the" _
    & " service name (long or registry format)","Extract SDDLText From Resource")
  End If 
Else
  Resource = wsh.arguments(0)
End If

If Resource = vbCancel then wsh.quit

If ucase(left(Resource,2)) = "HK" Then 
  ResType = "Registry"
  DebugMessage "is regkey"
  Hive = Left(Resource,instr(1,Resource,"\",1)-1)
  KeyName = Mid(Resource,instr(1,Resource,"\",1)+1)
  DebugMessage "Hive is: " & Hive
  DebugMessage "KeyName is: " & KeyName 
  SDDLText = ShowRegPerms(Hive, KeyName)
Elseif ucase(left(Resource,8)) = "SERVICE:" Then
  ResType = "Service"
  DebugMessage "is a service"
  ServiceName = Mid(Resource,9)
  SDDLText = ShowServicePerms(ServiceName)
Else
  ResType = "FileOrFolder"
  SDDLText = ShowFileorFolderPerms(Resource)
End If

If NOT (IsEmpty(SDDLText) OR IsNull(SDDLText)) Then
  if bInteractive then 
    inputbox "Press CTRL-C now to copy the SDDLText below: ","SDDLText Extractor",SDDLText
  else
    wsh.echo " ===> SDDLText below: " &vbcrlf &vbcrlf & SDDLText   
  End if 
Else
  wscript.echo "File, Folder or Registry Key May Not Exist"
End If

Function ShowServicePerms(ServiceName)
    ServiceWMIQuery = "Select * from Win32_Service WHERE Caption ='" & ServiceName & "' OR Name ='" & ServiceName & "'"
    DebugMessage "Service WMI query: " & ServiceWMIQuery
    Set colServices = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2").ExecQuery _
    (ServiceWMIQuery)
    
    debugmessage colServices.count
    
    If colServices.count = 1 Then 
      For Each objService in colServices
        debugmessage objService.Caption
        RetVal = objservice.GetSecurityDescriptor(wmiSecurityDescriptor)
        debugmessage "return from getsecuritydescriptor: " & RetVal
      Next
      sdHelper.Win32SDtoSDDL wmiSecurityDescriptor, returnSDDL
      ShowServicePerms = returnSDDL
    Else
      ShowServicePerms = ""
    End If
End Function

Function ShowRegPerms(RegHive,RegKey)
  oReg.GetSecurityDescriptor TLHiveValue(RegHive,"WMICONST"), RegKey, returnDACL
  sdHelper.Win32SDtoSDDL returnDACL, returnSDDL
  ShowRegPerms = returnSDDL
End Function

Function ShowFileorFolderPerms(FileorFolder)
  On Error Resume Next
  
  Set wmiFileSecSetting = GetObject( _
     "winmgmts:Win32_LogicalFileSecuritySetting.path='" & EscapeSlashes(FileorFolder) &"'")

  RetVal = wmiFileSecSetting.GetSecurityDescriptor(wmiSecurityDescriptor)
  If Err <> 0 Then
      On Error Goto 0
      Exit Function
  End If
  sdHelper.Win32SDtoSDDL wmiSecurityDescriptor, returnSDDL
  ShowFileorFolderPerms = returnSDDL 
  On Error Goto 0
End Function

Function TLHiveValue(FromHiveValue, ToHiveType)
  TLHiveType = "Value Not Found"
  
  Select Case Ucase(ToHiveType)
  Case "WMICONST"
    If FromHiveValue = WMI_HKLM OR FromHiveValue = "HKLM" OR FromHiveValue = "HKEY_LOCAL_MACHINE" Then TLHiveValue = WMI_HKLM 
    If FromHiveValue = WMI_HKCU OR FromHiveValue = "HKCU" OR FromHiveValue = "HKEY_CURRENT_USER" Then TLHiveValue = WMI_HKCU 
    If FromHiveValue = WMI_HKU OR FromHiveValue = "HKU" OR FromHiveValue = "HKEY_USERS" Then TLHiveValue = WMI_HKU 
  Case "SHORTTEXT"
    If FromHiveValue = "HKLM" OR FromHiveValue = WMI_HKLM OR FromHiveValue = "HKEY_LOCAL_MACHINE" Then TLHiveValue = "HKLM"
    If FromHiveValue = "HKCU" OR FromHiveValue = WMI_HKCU OR FromHiveValue = "HKEY_CURRENT_USER" Then TLHiveValue = "HKCU"
    If FromHiveValue = "HKU" OR FromHiveValue = WMI_HKU OR FromHiveValue = "HKEY_USERS" Then TLHiveValue = "HKU"
  Case "LONGTEXT"
    If FromHiveValue = "HKEY_LOCAL_MACHINE" OR FromHiveValue = WMI_HKLM OR FromHiveValue = "HKLM" Then TLHiveValue = "HKEY_LOCAL_MACHINE"
    If FromHiveValue = "HKEY_CURRENT_USER" OR FromHiveValue = WMI_HKCU OR FromHiveValue = "HKCU" Then TLHiveValue = "HKEY_CURRENT_USER"
    If FromHiveValue = "HKEY_USERS" OR FromHiveValue = WMI_HKU OR FromHiveValue = "HKU" Then TLHiveValue = "HKEY_USERS"
  Case Else
    TLHiveType = "Type Not Found" 
  End Select
  'DebugMessage "Translated: " & FromHiveValue & " to type: " & ToHiveType & ", Result: " & TLHiveValue

End Function

Function isAdmin()
  Const READ_CONTROL = &H20000
  Const HKEY_LOCAL_MACHINE = &H80000002
  
  Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  
  oReg.CheckAccess HKEY_LOCAL_MACHINE, "SECURITY", READ_CONTROL, isAdmin
  
End Function

Function EscapeSlashes(SourceString)
  EscapeSlashes = Replace(SourceString, "\", "\\")
End Function

Function DebugMessage (DbgMsg)
  MSGPREFIX = "DEBUG MSG: "
  If bDebugMessagesOn Then wsh.echo MSGPREFIX & DbgMsg
End Function

Function UserMessage (UsrMsg)
  wsh.echo UsrMsg
End Function
