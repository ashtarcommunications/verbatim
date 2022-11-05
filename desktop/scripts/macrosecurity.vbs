On Error Resume Next

Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

MsgBox "Ensure Word is closed, then click OK to set your macro security settings"

' All verisons of Office since 2016 have used the 16.0 version, but if that changes someday this will break
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Security\VBAWarnings", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Security\AccessVBOM", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Security\ProtectedView\DisableUnsafeLocationsInPv", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Security\ProtectedView\DisableInternetFilesInPV", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Security\ProtectedView\DisableAttachmentsInPV", 1, "REG_DWORD"