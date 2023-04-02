# Flow template

## Speech
*Send excel selection/cell/column to word at cursor or end of doc (`, Alt + `)
*Vtub for analytics

## Format on sheet
*Insert row(s) above/below current based on hotkey (make sure it doesn't screw up clipboard/undo) (F3, Alt + F3)
Insert cells above/below in same column (Ctrl + Alt + F3)
Delete row (Ctrl + F3)
Toggle enter/alt-enter mode
*Highlight argument/cell (F11)
Macro to designate evidence for an argument/cell (F7)
Paste special or expand by crlf (F6, Alt + F6)
Add/remove numbers for args automatically (Ctrl + Shift + 3)
Merge/group cells (F5)
Move up/down like with blocks (Ctrl + Alt + Up, Ctrl + Alt + Down)
*Keep f2 shortcut for entering cell 
Shortcut to extend to next column of same side (F9)

## Format workbook
Delete empty sheets
Delete single sheet
*Easy move between sheets
Add sheets quickly (aff/neg/etc), both hot key and ribbon menu (Ctrl + Alt + A, Ctrl + Alt + N)
*Label flows based on title in A1 (optionally disable)
Auto-move cursor to columns of chosen speech on all sheets (Ctrl + Alt + M)
Change row/column size
Add zoom functions to ribbon
*Configurable font size

## Misc features
Info tracking/scouting info
Automatically fill scouting info from subsheets
Auto-organize windows for flow/speech side-by-side
Settings form, configurable for different events
Keyboard shortcuts cheat sheet
Optional shorthand expansion

# New Tools
PC setup tool
Mac setup tool
	defaults read "Apple Global Domain" "com.apple.keyboard.fnState"
	defaults write "Apple Global Domain" "com.apple.keyboard.fnState" "1" ## F1 F2 etc
	defaults write "Apple Global Domain" "com.apple.keyboard.fnState" "0" ## Brightness/Media	https://apple.stackexchange.com/questions/344494/how-to-disable-default-mission-control-shortcuts-in-terminal
PC Installer
	Fix template getting put into admin templates folder when run with UAC (look at not needing admin permissions)
Plugin installer for Capture2Text, Timer, GetFromCiteCreator, NavPaneCycle, and EverythingSearch
Mac Installer
	Mac installer check/disbale Mac OS keyboard shortcuts

# Pre-release QA
Double check all Mac-specific macro modifications made it into the combined version
Check for extraneous styles
Test all functions on PC
Test all functions on Mac
Rubberduck everything
Run template through the decompiler
	
# Future ideas
Window arranger for multiple screens
Save all style customizations to settings to survive updates
Create flow for tabroom round/Speech from Verbatim
