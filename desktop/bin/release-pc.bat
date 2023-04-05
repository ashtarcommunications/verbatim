if "%1" == "" (
echo "Missing required release string parameter"
exit /b 1    
)
VBADecompiler.exe cmd/decompile /sXLfile:=..\Debate.dotm /sApplication:=Word /sAppVersion:=Default /bOverBackup:=0 /bPreservDateTime:=0 /bLogActivity:=0 /sFilePassword:=""
VBADecompiler.exe cmd/decompile /sXLfile:=..\DebateStartup.dotm /sApplication:=Word /sAppVersion:=Default /bOverBackup:=0 /bPreservDateTime:=0 /bLogActivity:=0 /sFilePassword:=""
VBADecompiler.exe cmd/decompile /sXLfile:=..\flow\Debate.xltm /sApplication:=Excel /sAppVersion:=Default /bOverBackup:=0 /bPreservDateTime:=0 /bLogActivity:=0 /sFilePassword:=""

del VBADecompiler.ini
del "..\Backup of Debate.dotm"
del "..\Backup of DebateStartup.dotm"
del "..\flow\Backup of Debate.xltm"

mkdir ..\release\%1
copy /y ..\Debate.dotm ..\release\%1\Debate.dotm
copy /y ..\DebateStartup.dotm ..\release\%1\DebateStartup.dotm
copy /y ..\flow\Debate.xltm ..\release\%1\Debate.xltm

"c:\Program Files (x86)\NSIS\makensis.exe" ..\install\Verbatim6.nsi
copy /y ..\install\Verbatim6.exe ..\release\%1\Verbatim6.exe

"c:\Program Files (x86)\NSIS\makensis.exe" ..\plugins\VerbatimPlugins.nsi
copy /y ..\..\timer\src-tauri\target\release\VerbatimTimer.exe ..\release\%1\VerbatimTimer.exe
copy /y ..\..\timer\src-tauri\target\release\bundle\msi\VerbatimTimer_1.0.0_x64_en-US.msi ..\release\%1\VerbatimTimer_1.0.0_x64_en-US.msi
copy /y ..\install\VerbatimPlugins.exe ..\release\%1\VerbatimPlugins.exe
copy /y ..\install\GetFromCiteCreator.exe ..\release\%1\GetFromCiteCreator.exe
copy /y ..\install\NavPaneCycle.exe ..\release\%1\NavPaneCycle.exe

copy /y "..\setup\Verbatim Setup Check\Bin\Release\net6.0-windows\win-x64\VerbatimSetupCheck.exe" ..\release\%1\VerbatimSetupCheck.exe
