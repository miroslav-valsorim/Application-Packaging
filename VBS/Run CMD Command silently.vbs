Set oaShell = CreateObject("WSCript.shell")
asd = oaShell.run("cmd /c del %systemdrive%\users\tests.txt /s", 0, true)
