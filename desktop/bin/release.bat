if "%1" == "" (
echo "Missing required version parameter (e.g. x.x.x)"
exit /b 1    
)

mkdir ..\release\%1
copy /y ..\Debate.dotm ..\release\%1\Debate.dotm
copy /y ..\DebateStartup.dotm ..\release\%1\DebateStartup.dotm
copy /y ..\flow\Debate.xltm ..\release\%1\Debate.xltm

copy /y ..\install\pc\Verbatim6.exe ..\release\%1\Verbatim6.exe
copy /y ..\install\mac\Verbatim6.pkg ..\release\%1\Verbatim6.pkg
copy /y ..\install\mac\VerbatimUninstall.zip ..\release\%1\VerbatimUninstall.zip

copy /y ..\plugins\VerbatimPlugins.exe ..\release\%1\VerbatimPlugins.exe

copy /y ..\..\timer\src-tauri\target\release\VerbatimTimer.exe ..\release\%1\VerbatimTimer.exe
copy /y ..\..\timer\release\VerbatimTimer_1.0.0_x64_en-US.msi ..\release\%1\VerbatimTimer_1.0.0_x64_en-US.msi
copy /y ..\..\timer\release\verbatim-timer_1.0.0_amd64.deb ..\release\%1\verbatim-timer_1.0.0_amd64.deb
copy /y ..\..\timer\release\VerbatimTimer_1.0.0_x64.dmg ..\release\%1\VerbatimTimer_1.0.0_x64.dmg

copy /y ..\plugins\GetFromCiteCreator.exe ..\release\%1\GetFromCiteCreator.exe
copy /y ..\plugins\NavPaneCycle.exe ..\release\%1\NavPaneCycle.exe

copy /y "..\setup\Verbatim Setup Check\Bin\Release\net6.0-windows\win-x64\publish\VerbatimSetupCheck.exe" ..\release\%1\VerbatimSetupCheck.exe
copy /y ..\setup\VerbatimSetupCheck.zip ..\release\%1\VerbatimSetupCheck.zip