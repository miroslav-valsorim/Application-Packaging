Dim FileName

sFileName = "Regional Settings Manager"
sDescription = "Regional Settings Manager"
sTargetPath = "c:\Program Files (x86)\Regional Settings Manager\RegionalSettingsMgr.exe"

Set shortcut = CreateObject("WScript.Shell").CreateShortcut(CreateObject("WScript.Shell").SpecialFolders("Startup") & + "\" + sFileName + ".lnk")
shortcut.Description = sDescription
shortcut.TargetPath = sTargetPath
shortcut.Save
