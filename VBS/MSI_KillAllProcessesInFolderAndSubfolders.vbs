' MSI script that kills every process started from the INSTALLDIR and it's subfolders
' Set the Custom Action after CostFinalize so the folders are defined

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * From Win32_Process")
Set oShell = CreateObject("WScript.Shell") 

' set folder name below ending with \ 
sINSTALLDIR = Session.Property("INSTALLDIR")

For Each objItem in colItems
    If Instr(1, objItem.ExecutablePath, sINSTALLDIR, vbTextCompare) > 0 Then objItem.terminate
Next