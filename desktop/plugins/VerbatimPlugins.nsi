; Verbatim Plugins Installer
; by Aaron Hardy
; v6.0.0
; 2023 Ashtar Communications

; nsUnzip plugin doesn't support Unicode
Unicode false

; Includes
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "nsProcess.nsh"
!include "Sections.nsh"

; The name of the installer
Name "Verbatim Plugins 6.0.0"

; The file to write
OutFile "VerbatimPlugins.exe"

; The default installation directory
InstallDir $PROGRAMFILES64\Verbatim\Plugins

; Configure UI
!define MUI_COMPONENTSPAGE_NODESC
!define MUI_ICON "..\assets\icons\Verbatim.ico"
!define MUI_UNICON "..\assets\icons\Verbatim.ico"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "..\assets\icons\Verbatim.bmp"

; Pages
!insertmacro MUI_PAGE_LICENSE "..\..\LICENSE"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
  
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

Section "Start Menu Shortcuts (Recommended)" sec1
  CreateDirectory "$SMPROGRAMS\Verbatim"
  CreateShortCut "$SMPROGRAMS\Verbatim\Verbatim Timer.lnk" "$PROGRAMFILES64\Verbatim\Plugins\VerbatimTimer.exe" "" "$PROGRAMFILES64\Verbatim\Plugins\VerbatimTimer.exe" 0
  CreateShortCut "$SMPROGRAMS\Verbatim\Capture2Text.lnk" "$PROGRAMFILES64\Verbatim\Plugins\OCR\Capture2Text.exe" "" "$PROGRAMFILES64\Verbatim\Plugins\OCR\Capture2Text.exe" 0
  CreateShortCut "$SMPROGRAMS\Verbatim\Everything Search.lnk" "$PROGRAMFILES64\Verbatim\Plugins\Search\Everything.exe" "" "$PROGRAMFILES64\Verbatim\Plugins\Search\Everything.exe" 0
  CreateShortCut "$SMPROGRAMS\Verbatim\Uninstall Verbatim Plugins.lnk" "$INSTDIR\UninstallVerbatimPlugins.exe" "" "$INSTDIR\UninstallVerbatimPlugins.exe" 0
SectionEnd

Section "Verbatim Timer (Required)" sec2
	; Make mandatory
	SectionIn RO

	; Set output path to the installation directory.
	SetOutPath $INSTDIR

	File "..\..\timer\src-tauri\target\release\VerbatimTimer.exe"
	
	; Write Uninstall registry keys and file
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "DisplayName" "Verbatim Plugins"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "Publisher" "Ashtar Communications"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "HelpLink" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "URLUpdateInfo" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "URLInfoAbout" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "DisplayVersion" "6.0.0"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "UninstallString" '"$INSTDIR\UninstallVerbatimPlugins.exe"'
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "NoModify" 1
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins" "NoRepair" 1
	WriteUninstaller "UninstallVerbatimPlugins.exe"
SectionEnd

Section "Capture2Text OCR" sec3
	CreateDirectory "$PROGRAMFILES64\Verbatim\Plugins\OCR"
	SetOutPath $PROGRAMFILES64\Verbatim\Plugins\OCR
	File "ocr\Capture2Text_v4.6.3_64bit.zip"
	nsUnzip::Extract "Capture2Text_v4.6.3_64bit.zip" /u /END
SectionEnd

Section "Everything Search" sec4
    CreateDirectory "$PROGRAMFILES64\Verbatim\Plugins\Search"
	SetOutPath $PROGRAMFILES64\Verbatim\Plugins\Search
	File "search\Everything.exe"
	File "search\Everything.lng"
	File "search\Everything-license.txt"
SectionEnd

Section "NavPaneCycle" sec5
	SetOutPath $INSTDIR
	File "NavPaneCycle.exe"
SectionEnd

Section "GetFromCiteCreator" sec6
	SetOutPath $INSTDIR
	File "GetFromCiteCreator.exe"
SectionEnd

Section "Uninstall"
    ; Kill plugins if running
	${nsProcess::KillProcess} "VerbatimTimer.exe" $R0
	${nsProcess::KillProcess} "Capture2Text.exe" $R0
	${nsProcess::KillProcess} "Capture2Text_CLI.exe" $R0
	${nsProcess::KillProcess} "Everything.exe" $R0
	
	; Remove plugins directory
	RMDir /r "$INSTDIR"

	; Remove shortcuts, if any
	Delete "$SMPROGRAMS\Verbatim\Verbatim Timer*"
	Delete "$SMPROGRAMS\Verbatim\Capture2Text*"
	Delete "$SMPROGRAMS\Verbatim\Everything*"
	Delete "$SMPROGRAMS\Verbatim\Uninstall Verbatim Plugins*"
	  
	; Remove installer registry keys
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimPlugins"
	
	${nsProcess::Unload}
SectionEnd
