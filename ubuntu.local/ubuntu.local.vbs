Option Explicit
Const vbHide = 0
Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
objWShell.Run "C:\Progra~1\Oracle\VirtualBox\VBoxHeadless.exe --startvm ubuntu.local", vbHide
Set objWShell = Nothing