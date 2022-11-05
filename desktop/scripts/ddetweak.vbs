On Error Resume Next
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
MsgBox "Click OK to change your DDE settings - this will not take effect until you restart Word."
WshShell.RegWrite "HKEY_CLASSES_ROOT\Word.Document.8\shell\Open\ddeexec.backup\", "", "REG_SZ"
WshShell.RegWrite "HKEY_CLASSES_ROOT\Word.Document.12\shell\Open\ddeexec.backup\", "", "REG_SZ"
WshShell.RegWrite "HKEY_CLASSES_ROOT\Word.Document.8\shell\Open\ddeexec\", "[FileOpen(""%1"")]", "REG_SZ"
WshShell.RegWrite "HKEY_CLASSES_ROOT\Word.Document.12\shell\Open\ddeexec\", "[FileOpen(""%1"")]", "REG_SZ"
WScript.Quit
