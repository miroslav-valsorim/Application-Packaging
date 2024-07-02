'==========================================================================
'
' DESCRIPTION: Deletes Reg Key
'
' NAME: DelRegKeys.vbs
'
'
' USAGE: Deletes Reg Key and all SubKeys and Values.
'
' PREREQ: WMI must be installed on host machine. Use only if all other 
'         methods of deleting registry keys do not work.
'
' COMMENT: Msi has this functionality already.
'
'==========================================================================
Option Explicit

Const HKCR = &H80000000
Const HKCU = &H80000001
Const HKLM = &H80000002
Const HKU  = &H80000003
Const HKCC = &H80000005

Dim objWMI, objServer, objReg
Dim arrRemoveKeys, strRemoveKey


	'Keys to be deleted. Use HK** format (i.e. HKCU\SOFTWARE\Microsoft).
	
arrRemoveKeys = Array( _
"HKLM\SOFTWARE\GATE" _
)

Set objWMI = CreateObject("wBEMScripting.sWBEMLocator")
Set objServer = objWMI.ConnectServer(".","root\default")
Set objReg = objServer.Get("StdRegProv")

For Each strRemoveKey In arrRemoveKeys
	RemoveKey strRemoveKey
Next

Set objReg = Nothing
Set objServer = Nothing
Set objWMI = Nothing


Sub RemoveKey(strKey)
	Dim arrKey, strRoot, i

	' Remove trailing backslash
	If (Right(strKey, 1) = "\") Then
		strKey = Left(strKey, (Len(strKey) - 1))
	End If

	' Determining registry root hive
	arrKey = Split(strKey, "\")
	If (UBound(arrKey) < 2) Then
		Exit Sub
	End If
	Select Case arrKey(0)
		Case "HKCR"
			strRoot = HKCR
		Case "HKCU"
			strRoot = HKCU
		Case "HKLM"
			strRoot = HKLM
		Case "HKU"
			strRoot = HKU
		Case "HKCC"
			strRoot = HKCC
	End Select

	' Determining main registry key
	strKey = arrKey(1)
	For i = 2 To (UBound(arrKey) - 1)
		strKey = strKey & "\" & arrKey(i)
	Next

	' Remove subkey
	RecurseKeys strRoot, strKey, arrKey(UBound(arrKey))
End Sub

Sub RecurseKeys(strRoot, strKey, strSubKey)
	Dim strFullKey, arrSubKeys, arrValues, arrValueTypes, strValue

	strFullKey = strKey & "\" & strSubKey

	' Enumerate subkeys
	objReg.EnumKey strRoot, strFullKey, arrSubKeys
	If (IsNull(arrSubKeys) = False) Then
		For Each strSubKey In arrSubKeys
			RecurseKeys strRoot, strFullKey, strSubKey
		Next
	End If

	' Enumerate and delete values
	objReg.EnumValues strRoot, strFullKey, arrValues, arrValueTypes
	If (IsNull(arrValues) = False) Then
		For Each strValue In arrValues
			objReg.DeleteValue strRoot, strFullKey, strValue
		Next
	End If

	' Delete key
	objReg.DeleteKey strRoot, strFullKey
End Sub
