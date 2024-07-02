'************************************************************************** * 
' Version: 1.0
' Description: Reads the language value of Windows OS system locale. Stores
' 			   the result in Property within the MSI (string).
' You can then use this property for multiple language setups.
'************************************************************************** * 

Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
for each System in SystemSet 

' Store System Language Locale (will use HEX):
 Hex_loc = System.Locale
' Convert to decimal
 Dec_loc = CInt("&h" & Hex_loc)
' Convert decimal to string
 strLOCALE = CStr(Dec_loc)
' Below is a test line:
' WScript.Echo " Locale: " + strLOCALE
 
Session.Property("HP_LANG")=strLOCALE
 
next