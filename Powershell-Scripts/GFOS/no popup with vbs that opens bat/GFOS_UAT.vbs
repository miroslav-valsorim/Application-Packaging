Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "C:\Program Files (x86)\GraphOn\AppController\GFOS_UAT.bat" & Chr(34), 0
Set WshShell = Nothing