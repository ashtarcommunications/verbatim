On Error Resume Next
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
MsgBox "Click OK to clear your DDE settings - this will not take effect until you restart Word."
WScript.Quit
