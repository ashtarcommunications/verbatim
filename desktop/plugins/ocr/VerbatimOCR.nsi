;Verbatim OCR Installer
;by Aaron Hardy
;v1.0
;2022-11-23

Unicode False

; Includes
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "nsProcess.nsh"
!include "Sections.nsh"

; The name of the installer
Name "Verbatim OCR 1.0"

; The file to write
OutFile "VerbatimOCR.exe"

; The default installation directory
InstallDir $PROGRAMFILES64\Verbatim\OCR

; Configure UI
!define MUI_COMPONENTSPAGE_NODESC
!define MUI_ICON "Verbatim.ico"
!define MUI_UNICON "Verbatim.ico"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "Verbatim.bmp"

; Pages
!insertmacro MUI_PAGE_LICENSE "license.txt"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
  
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

Section "Start Menu Shortcuts (Recommended)" sec1
  CreateDirectory "$SMPROGRAMS\Verbatim"
  CreateShortCut "$SMPROGRAMS\Verbatim\Capture2Text.lnk" "$PROGRAMFILES64\Verbatim\OCR\Capture2Text.exe" "" "$PROGRAMFILES64\Verbatim\OCR\Capture2Text.exe" 0
  CreateShortCut "$SMPROGRAMS\Verbatim\Uninstall Verbatim OCR.lnk" "$INSTDIR\UninstallVerbatimOCR.exe" "" "$INSTDIR\UninstallVerbatimOCR.exe" 0
SectionEnd

Section "Capture2Text OCR (Required)" sec6

	; Make mandatory
	SectionIn RO
  
	; Set output path to the installation directory.
	SetOutPath $INSTDIR
  
	; Put files there
	File "Capture2Text_v4.6.3_64bit.zip"
	nsUnzip::Extract "Capture2Text_v4.6.3_64bit.zip" /u /END
	
	;Write Uninstall registry keys and file
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimOCR" "DisplayName" "VerbatimOCR"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "Publisher" "Ashtar Communications"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "HelpLink" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "URLUpdateInfo" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "URLInfoAbout" "https://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "DisplayVersion" "1.0"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "UninstallString" '"$INSTDIR\UninstallVerbatimOCR.exe"'
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "NoModify" 1
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "NoRepair" 1
	WriteUninstaller "UninstallVerbatimOCR.exe"
	
SectionEnd

Section "Uninstall"
  
    ;Kill Capture2Text if it's running
	${nsProcess::KillProcess} "Capture2Text.exe" $RO
	${nsProcess::KillProcess} "Capture2Text_CLI.exe" $RO
	
	; Remove files and uninstaller
	;Delete "$INSTDIR\UninstallVerbatimOCR.exe"
	RMDir /r "$INSTDIR"

	; Remove shortcuts, if any
	Delete "$SMPROGRAMS\Verbatim\Capture2Text*"
	  
	; Remove installer registry keys
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\VerbatimOCR"
	
	${nsProcess::Unload}
SectionEnd
