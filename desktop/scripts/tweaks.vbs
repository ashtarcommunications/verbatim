On Error Resume Next
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
MsgBox "Click OK to change tweaks."
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShowPreviewHandlers", "0", "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\Graphics\DisableHardwareAcceleration", "1", "REG_DWORD"
WScript.Quit