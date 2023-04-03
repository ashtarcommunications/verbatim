;AHK to fetch current cite in Cite Creator chrome extension
;Not persistent - designed to be called from Word VBA
;Version 3.0.0
;2023 Ashtar Communications

#SingleInstance force

;If Chrome is open, activate it with the current tab
if WinExist("ahk_class Chrome_WidgetWin_1")
{
	;Clear clipboard, press Cite Maker shortcut (Ctrl+Alt+C), and wait for clipboard to update
	clipboard =
    WinActivate
    Send ^!c
	ClipWait, 2
	
	;If nothing copied, exit
	if ErrorLevel
	{
		return
	}

	;Get last character as an ASCII code
	StringRight, lastchar, clipboard, 1
	Transform, lastchar, Asc, %lastchar%

	;Count number of newlines
	StringReplace, clipboard, clipboard, `r`n, `r`n, UseErrorLevel
	numlines = %ErrorLevel%
	
	;Switch back to Word, paste unformatted and run auto-format cite
	if WinExist("ahk_class OpusApp")
	{
		WinActivate
		Send {F2}

		;Move to the start of the cite line
		if numlines = 0
		{
			Send ^{Up}
		}
		else if numlines = 1
		{
			Send ^{Up}
			if lastchar <> 10
				Send ^{Up}	
		}
		else 
		{			
			Loop %numlines% {
				Send ^{Up}
			}
			if lastchar <> 10
				Send ^{Up}	
		}
		
		Send ^{F8}
	}
}
return