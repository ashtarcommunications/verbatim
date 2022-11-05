;Verbatim Installer
;by Aaron Hardy
;v5.1
;10-2-2014

; Includes
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "nsProcess.nsh"
!include "registry.nsh"
!include "Sections.nsh"

; The name of the installer
Name "Verbatim 5.1"

; The file to write
OutFile "Verbatim5.exe"

; The default installation directory
InstallDir $APPDATA\Microsoft\Templates

; Request application privileges - admin required for registry tweaks, so "highest" will give them it possible
RequestExecutionLevel highest

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

;Insert Copy Registry Key Macro
!insertmacro COPY_REGISTRY_KEY
!insertmacro UN.COPY_REGISTRY_KEY

Section "Start Menu Shortcuts (Recommended)" sec1
  CreateDirectory "$SMPROGRAMS\Verbatim"
  CreateShortCut "$SMPROGRAMS\Verbatim\Verbatim.lnk" "$APPDATA\Microsoft\Templates\Debate.dotm" "" "$APPDATA\Microsoft\Templates\Debate.dotm" 0
  CreateShortCut "$SMPROGRAMS\Verbatim\Uninstall Verbatim.lnk" "$INSTDIR\UninstallVerbatim.exe" "" "$INSTDIR\UninstallVerbatim.exe" 0
SectionEnd

Section "Desktop Shortcut (Recommended)" sec2
	CreateShortCut "$DESKTOP\Verbatim.lnk" "$APPDATA\Microsoft\Templates\Debate.dotm" "" "$APPDATA\Microsoft\Templates\Debate.dotm" 0
SectionEnd

Section "Additional Tools (Recommended)" sec3
	SetOutPath $INSTDIR
	File "Timer.exe" 
	File "NavPaneCycle.exe"
	File "GetFromCiteMaker.exe"
SectionEnd

Section "OCR Support (Recommended)" sec4
	SetOutPath $INSTDIR
	File /r "Capture2Text"    
SectionEnd

SectionGroup "Registry Tweaks (Recommended)" sec5
	Section "Set Macro Security Level" sec5a
		;Set for Office 2013
		ReadRegDWORD $0 HKCU Software\Microsoft\Office\15.0\Word\Security VBAWarnings
		${If} $0 > 1
			WriteRegDWORD HKCU "Software\Microsoft\Office\15.0\Word\Security" "VBAWarnings" 1
		${EndIf}
		
		ReadRegDWORD $1 HKCU Software\Microsoft\Office\15.0\Word\Security AccessVBOM
		${If} $1 == 0
			WriteRegDWORD HKCU "Software\Microsoft\Office\15.0\Word\Security" "AccessVBOM" 1
		${EndIf}
		
		;Set for Office 2010
		ReadRegDWORD $2 HKCU Software\Microsoft\Office\14.0\Word\Security VBAWarnings
		${If} $2 > 1
			WriteRegDWORD HKCU "Software\Microsoft\Office\14.0\Word\Security" "VBAWarnings" 1
		${EndIf}
		
		ReadRegDWORD $3 HKCU Software\Microsoft\Office\14.0\Word\Security AccessVBOM
		${If} $3 == 0
			WriteRegDWORD HKCU "Software\Microsoft\Office\14.0\Word\Security" "AccessVBOM" 1
		${EndIf}
		
	SectionEnd
	
	Section "Disable Protected View" sec5b
	
		;Set for Office 2013
		ReadRegDWORD $0 HKCU Software\Microsoft\Office\15.0\Word\Security\ProtectedView DisableUnsafeLocationsInPV
		${If} $0 == 0
			WriteRegDWORD HKCU "Software\Microsoft\Office\15.0\Word\Security\ProtectedView" "DisableUnsafeLocationsInPV" 1
		${EndIf}
		
		ReadRegDWORD $1 HKCU Software\Microsoft\Office\15.0\Word\Security\ProtectedView DisableInternetFilesInPV
		${If} $1 == 0
			WriteRegDWORD HKCU "Software\Microsoft\Office\15.0\Word\Security\ProtectedView" "DisableInternetFilesInPV" 1
		${EndIf}
		
		ReadRegDWORD $2 HKCU Software\Microsoft\Office\15.0\Word\Security\ProtectedView DisableAttachmentsInPV
		${If} $2 == 0
			WriteRegDWORD HKCU "Software\Microsoft\Office\15.0\Word\Security\ProtectedView" "DisableAttachmentsInPV" 1
		${EndIf}
		
		;Set for Office 2010	
		ReadRegDWORD $3 HKCU Software\Microsoft\Office\14.0\Word\Security\ProtectedView DisableUnsafeLocationsInPV
		${If} $3 == 0
			WriteRegDWORD HKCU "Software\Microsoft\Office\14.0\Word\Security\ProtectedView" "DisableUnsafeLocationsInPV" 1
		${EndIf}
		
		ReadRegDWORD $4 HKCU Software\Microsoft\Office\14.0\Word\Security\ProtectedView DisableInternetFilesInPV
		${If} $4 == 0
			WriteRegDWORD HKCU "Software\Microsoft\Office\14.0\Word\Security\ProtectedView" "DisableInternetFilesInPV" 1
		${EndIf}
		
		ReadRegDWORD $5 HKCU Software\Microsoft\Office\14.0\Word\Security\ProtectedView DisableAttachmentsInPV
		${If} $5 == 0
			WriteRegDWORD HKCU "Software\Microsoft\Office\14.0\Word\Security\ProtectedView" "DisableAttachmentsInPV" 1
		${EndIf}
		
	SectionEnd
	
	Section "Set DDE to Single Instance" sec5c
		
		;Make a backup of current DDE key if it exists
		${COPY_REGISTRY_KEY} HKCR "Word.Document.12\shell\Open\ddeexec" HKCR "Word.Document.12\shell\Open\ddeexec.backup"
		${COPY_REGISTRY_KEY} HKCR "Word.Document.8\shell\Open\ddeexec" HKCR "Word.Document.8\shell\Open\ddeexec.backup"
	
		;May only work on Word 2013 - check on Word 2010
		WriteRegStr HKCR "Word.Document.12\shell\Open\ddeexec" "" "[FileOpen($\"%1$\")]"
		WriteRegStr HKCR "Word.Document.8\shell\Open\ddeexec" "" "[FileOpen($\"%1$\")]"
	
	SectionEnd
	
	Section "Disable Hardware Acceleration" sec5d
		
		;Test if Office 2013 is installed
		ClearErrors
		EnumRegKey $0 HKCU "Software\Microsoft\Office\15.0\Common" 0
		IfErrors 0 keyexist
			Goto end ;Else skip and goto end
		keyexist:
			WriteRegDWORD HKCU "Software\Microsoft\Office\15.0\Common\Graphics" "DisableHardwareAcceleration" 1
		
		end:
		ClearErrors
	SectionEnd
	
	Section "Disable Explorer Preview Pane" sec5e
			WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowPreviewHandlers" 0
	SectionEnd
SectionGroupEnd

Section "Verbatim (Required)" sec6

	; Make mandatory
	SectionIn RO
  
	; Set output path to the installation directory.
	SetOutPath $INSTDIR
  
	; Put files there
	File "Debate.dotm"
	File "changelog.txt"
	
	; Remove old registry keys if pre v. 5, otherwise reset FirstRun
	ReadRegStr $0 HKCU "Software\VB And VBA Program Settings\Verbatim\Main" Version
	${If} $0 >= "5"
		WriteRegStr HKCU "Software\VB And VBA Program Settings\Verbatim\Main" "Version" "5.1"
		WriteRegStr HKCU "Software\VB And VBA Program Settings\Verbatim\Admin" "FirstRun" "True"
	${Else}
		DeleteRegKey HKCU "Software\VB And VBA Program Settings\Verbatim"	
	${EndIf}
  	
	;Write Uninstall registry keys and file
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "DisplayName" "Verbatim"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "Publisher" "Ashtar Communications"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "HelpLink" "http://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "URLUpdateInfo" "http://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "URLInfoAbout" "http://paperlessdebate.com"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "DisplayVersion" "5.1"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "UninstallString" '"$INSTDIR\UninstallVerbatim.exe"'
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "NoModify" 1
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim" "NoRepair" 1
	WriteUninstaller "UninstallVerbatim.exe"
	
SectionEnd

Section "Uninstall"
  
	${nsProcess::KillProcess} "Timer.exe" $R0	;Kill Timer if it's running
	  
	;Check if Word is running
	retry:
	${nsProcess::FindProcess} "Winword.exe" $R0

	${If} $R0 == 0
		MessageBox MB_ABORTRETRYIGNORE|MB_DEFBUTTON2 "Word appears to be running. Please ensure it is completely closed before proceeding. $\n$\n After closing Word, click Retry to proceed. $\n$\n If you click Ignore, Word will be force quit without saving." IDRETRY retry IDIGNORE ignore
	${Else}
		Goto skip
	${EndIf}

	;cancel:
	Abort "Uninstall cancelled - Word was still running."

	ignore:
	MessageBox MB_YESNO "Are you sure you want to force quit Word?" IDYES quitword IDNO retry

	quitword:
	${nsProcess::KillProcess} "Winword.exe" $R0

	skip:
	${nsProcess::Unload}
  
	; Remove files and uninstaller
	Delete "$INSTDIR\Debate.dotm"
	Delete "$INSTDIR\changelog.txt"
	Delete "$INSTDIR\Timer.exe"
	Delete "$INSTDIR\NavPaneCycle.exe"
	Delete "$INSTDIR\GetFromCiteMaker.exe"
	RMDir /r "$INSTDIR\Capture2Text"
	Delete "$INSTDIR\UninstallVerbatim.exe"

	; Remove shortcuts, if any
	Delete "$SMPROGRAMS\Verbatim\*.*"
	RMDir "$SMPROGRAMS\Verbatim"
	Delete "$DESKTOP\Verbatim.lnk"
	  
	;IF DDE backup was created, restore it
	ClearErrors
	EnumRegKey $0 HKCR "Word.Document.12\shell\Open\ddeexec.backup" 0
	IfErrors 0 keyexist
		Goto end ;Else skip and goto end
	keyexist:
		DeleteRegKey HKCR "Word.Document.12\shell\Open\ddeexec"
		${UN.COPY_REGISTRY_KEY} HKCR "Word.Document.12\shell\Open\ddeexec.backup" HKCR "Word.Document.12\shell\Open\ddeexec"
		DeleteRegKey HKCR "Word.Document.12\shell\Open\ddeexec.backup"
			
		DeleteRegKey HKCR "Word.Document.8\shell\Open\ddeexec"
		${UN.COPY_REGISTRY_KEY} HKCR "Word.Document.8\shell\Open\ddeexec.backup" HKCR "Word.Document.8\shell\Open\ddeexec"
		DeleteRegKey HKCR "Word.Document.8\shell\Open\ddeexec.backup"
	
	end:
	ClearErrors
	
	; Remove installer registry keys
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Verbatim"
	
SectionEnd

; On initialization, check whether Word is running
Function .onInit

	; Check for admin rights and hide/disable registry section if not admin
	UserInfo::getAccountType
	Pop $0
	${If} $0 != "Admin"
		!insertmacro UnselectSection ${sec5}
		SectionSetText ${sec5} ""
		SectionSetText ${sec5a} ""
		SectionSetText ${sec5b} ""
		SectionSetText ${sec5c} ""
		SectionSetText ${sec5d} ""
		SectionSetText ${sec5e} ""
	${EndIf}
		
	retry:
	${nsProcess::KillProcess} "Timer.exe" $R0	;Kill Timer if it's running
	${nsProcess::FindProcess} "Winword.exe" $R0
	
	${If} $R0 == 0
		MessageBox MB_ABORTRETRYIGNORE|MB_DEFBUTTON2 "Word appears to be running. Please ensure it is completely closed before proceeding. $\n$\n After closing Word, click Retry to proceed. $\n$\n If you click Ignore, Word will be force quit without saving." IDRETRY retry IDIGNORE ignore
	${Else}
		Goto end
	${EndIf}

	;cancel:
	Abort "Installation cancelled - Word was still running."
	
	ignore:
	MessageBox MB_YESNO "Are you sure you want to force quit Word?" IDYES quitword IDNO retry
	
	quitword:
	${nsProcess::KillProcess} "Winword.exe" $R0
	
	end:
	${nsProcess::Unload}
FunctionEnd
