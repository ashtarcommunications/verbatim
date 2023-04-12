VBADecompiler.exe cmd/decompile /sXLfile:=..\Debate.dotm /sApplication:=Word /sAppVersion:=Default /bOverBackup:=0 /bPreservDateTime:=0 /bLogActivity:=0 /sFilePassword:=""
VBADecompiler.exe cmd/decompile /sXLfile:=..\DebateStartup.dotm /sApplication:=Word /sAppVersion:=Default /bOverBackup:=0 /bPreservDateTime:=0 /bLogActivity:=0 /sFilePassword:=""
VBADecompiler.exe cmd/decompile /sXLfile:=..\flow\Debate.xltm /sApplication:=Excel /sAppVersion:=Default /bOverBackup:=0 /bPreservDateTime:=0 /bLogActivity:=0 /sFilePassword:=""

del VBADecompiler.ini
del "..\Backup of Debate.dotm"
del "..\Backup of DebateStartup.dotm"
del "..\flow\Backup of Debate.xltm"

"c:\Program Files (x86)\NSIS\makensis.exe" ..\install\pc\Verbatim6.nsi

"c:\Program Files (x86)\NSIS\makensis.exe" ..\plugins\VerbatimPlugins.nsi
