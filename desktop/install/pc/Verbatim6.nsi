; Verbatim Installer
; by Aaron Hardy
; v6.0.0
; 2023 Ashtar Communications

Unicode true

; Includes
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "nsProcess.nsh"
!include "registry.nsh"
!include "Sections.nsh"

; The name of the installer
Name "Verbatim 6.0.0"

; The file to write
OutFile "Verbatim6.exe"

; The default installation directory
InstallDir $APPDATA\Microsoft\Templates

; Configure UI
!define MUI_COMPONENTSPAGE_NODESC
!define MUI_ICON "..\..\assets\icons\Verbatim.ico"
!define MUI_UNICON "..\..\assets\icons\Verbatim.ico"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "..\..\assets\icons\Verbatim.bmp"

; Pages
!insertmacro MUI_PAGE_LICENSE "..\..\..\LICENSE"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
  
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

Section "Start Menu Shortcuts (Recommended)" sec1
  CreateDirectory "$SMPROGRAMS\Verbatim"
  CreateShortCut "$SMPROGRAMS\Verbatim\Verbatim.lnk" "$APPDATA\Microsoft\Templates\Debate.dotm" "" "$APPDATA\Microsoft\Templates\Debate.dotm" 0
  CreateShortCut "$SMPROGRAMS\Verbatim\Verbatim Flow.lnk" "$APPDATA\Microsoft\Templates\Debate.xltm" "" "$APPDATA\Microsoft\Templates\Debate.xltm" 0
  CreateShortCut "$SMPROGRAMS\Verbatim\Uninstall Verbatim.lnk" "$PROGRAMFILES64\Verbatim\UninstallVerbatim.exe" "" "$PROGRAMFILES64\Verbatim\UninstallVerbatim.exe" 0
SectionEnd

Section "Desktop Shortcut (Recommended)" sec2
	CreateShortCut "$DESKTOP\Verbatim.lnk" "$APPDATA\Microsoft\Templates\Debate.dotm" "" "$APPDATA\Microsoft\Templates\Debate.dotm" 0
	CreateShortCut "$DESKTOP\Verbatim Flow.lnk" "$APPDATA\Microsoft\Templates\Debate.xltm" "" "$APPDATA\Microsoft\Templates\Debate.xltm" 0
SectionEnd

Section "Verbatim (Required)" sec3
	; Make mandatory
	SectionIn RO
  
	; Create Program Files directories
	CreateDirectory "$PROGRAMFILES64\Verbatim"
	CreateDirectory "$PROGRAMFILES64\Verbatim\Plugins"
	CreateShortCut "$PROGRAMFILES64\Verbatim\Verbatim.lnk" "$APPDATA\Microsoft\Templates\Debate.dotm" "" "$APPDATA\Microsoft\Templates\Debate.dotm" 0
	CreateShortCut "$PROGRAMFILES64\Verbatim\Verbatim Flow.lnk" "$APPDATA\Microsoft\Templates\Debate.xltm" "" "$APPDATA\Microsoft\Templates\Debate.xltm" 0
  
	; Set output path to the installation directory.
	SetOutPath $INSTDIR
  
	; Put template files there
	File "..\..\Debate.dotm"
	File "..\..\flow\Debate.xltm"

	; Clean up old installs
	Delete "$INSTDIR\Timer.exe"
	Delete "$INSTDIR\changelog.txt"
	Delete "$INSTDIR\NavPaneCycle.exe"
	Delete "$INSTDIR\GetFromCiteMaker.exe"
	RMDir /r "$INSTDIR\Capture2Text"
	
	; Install startup file
	SetOutPath $APPDATA\Microsoft\Word\STARTUP
	File "..\..\DebateStartup.dotm"
	
	; Install changelog
	SetOutPath $PROGRAMFILES64\Verbatim
	File "..\..\CHANGELOG.md"
	
	; Remove old registry keys pre v6, otherwise reset FirstRun
	ReadRegStr $0 HKCU "Software\VB And VBA Program Settings\Verbatim\Profile" Version
	${If} $0 >= "5"
		WriteRegStr HKCU "Software\VB And VBA Program Settings\Verbatim\Profile" "Version" "6.0.0"
		WriteRegStr HKCU "Software\VB And VBA Program Settings\Verbatim\Admin" "FirstRun" "True"
	${Else}
		DeleteRegKey HKCU "Software\VB And VBA Program Settings\Verbatim"	
	${EndIf}
  	
	; Write Uninstall registry keys and file
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "DisplayName" "Verbatim"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "Publisher" "Ashtar Communications"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "HelpLink" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "URLUpdateInfo" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "URLInfoAbout" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "DisplayVersion" "6.0.0"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "UninstallString" '"$PROGRAMFILES64\Verbatim\UninstallVerbatim.exe"'
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "NoModify" 1
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "NoRepair" 1
	WriteUninstaller "$PROGRAMFILES64\Verbatim\UninstallVerbatim.exe"	
SectionEnd

Section "Uninstall"
	; Kill plugins, since entire Verbatim directory is deleted
	${nsProcess::KillProcess} "VerbatimTimer.exe" $R0
	${nsProcess::KillProcess} "Capture2Text.exe" $R0
	${nsProcess::KillProcess} "Capture2Text_CLI.exe" $R0
	${nsProcess::KillProcess} "Everything.exe" $R0

  	; Check if Word is running
	retry:
	${nsProcess::FindProcess} "Winword.exe" $R0

	${If} $R0 == 0
		MessageBox MB_ABORTRETRYIGNORE|MB_DEFBUTTON2 "Word appears to be running. Please ensure it is completely closed before proceeding. $\n$\n After closing Word, click Retry to proceed. $\n$\n If you click Ignore, Word will be force quit without saving." IDRETRY retry IDIGNORE ignore
	${Else}
		Goto skip
	${EndIf}

	; cancel:
	Abort "Uninstall cancelled - Word was still running."

	ignore:
	MessageBox MB_YESNO "Are you sure you want to force quit Word?" IDYES quitword IDNO retry

	quitword:
	${nsProcess::KillProcess} "Winword.exe" $R0

	skip:
	${nsProcess::Unload}
  
	; Remove files and uninstaller
	Delete "$APPDATA\Microsoft\Templates\Debate.dotm"
	Delete "$APPDATA\Microsoft\Templates\Debate.xltm"
	Delete "$APPDATA\Microsoft\Templates\DebateAnalytics.xlsx"
	Delete "$APPDATA\Microsoft\Word\STARTUP\DebateStartup.dotm"

	; Remove shortcuts, if any
	Delete "$SMPROGRAMS\Verbatim\*.*"
	RMDir /r "$SMPROGRAMS\Verbatim"
	Delete "$PROGRAMFILES64\Verbatim\*.*"
	RMDir /r "$PROGRAMFILES64\Verbatim"
	Delete "$DESKTOP\Verbatim.lnk"
	Delete "$DESKTOP\Verbatim Flow.lnk"
	  
	; Remove installer registry keys
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim"
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins"
SectionEnd

; On initialization, check whether Word is running
Function .onInit
	retry:
	${nsProcess::FindProcess} "Winword.exe" $R0
	
	${If} $R0 == 0
		MessageBox MB_ABORTRETRYIGNORE|MB_DEFBUTTON2 "Word appears to be running. Please ensure it is completely closed before proceeding. $\n$\n After closing Word, click Retry to proceed. $\n$\n If you click Ignore, Word will be force quit without saving." IDRETRY retry IDIGNORE ignore
	${Else}
		Goto end
	${EndIf}

	; cancel:
	Abort "Installation cancelled - Word was still running."
	
	ignore:
	MessageBox MB_YESNO "Are you sure you want to force quit Word?" IDYES quitword IDNO retry
	
	quitword:
	${nsProcess::KillProcess} "Winword.exe" $R0
	
	end:
	${nsProcess::Unload}
FunctionEnd
