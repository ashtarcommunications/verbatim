# Bugs & New Features
Search plugin
Troubleshooting form
Tutorial Replacement, consider a sample card in building blocks
Rewrite Find blocks to use a range where possible
Unify card/doc selection modes (shrink, condense, uncondense, shrink pilcrows, cite-ify, convert dates, etc.)

# Flow template
Send selected or heading from Word to excel in cell or column
Send outline tag/cite from word to excel
Send excel selection/cell/column to word at cursor or end of doc
Insert row(s) above/below current based on hotkey (make sure it doesn't screw up clipboard/undo)
Insert cells above/below in same column
Delete row
Delete empty sheets
Delete single sheet
Easy move between sheets
Add sheets quickly (aff/neg/etc), both hot key and ribbon menu
Toggle enter/alt-enter mode
Highlight argument/cell
Create flow for tabroom round/Speech from Verbatim
Label flows based on title in A1 (optionally disable)
Info tracking/scouting info
Auto-move cursor to columns of chosen speech on all sheets
Macro to designate evidence for an argument/cell
Auto-organize windows for flow/speech side-by-side
Paste special or expand by crlf
Vtub for analytics
Add/remove numbers for args automatically
Merge/group cells
Move up/down like with blocks
Configurable for different events
Keyboard shortcuts cheat sheet
Keep f2 shortcut for entering cell
Change row/column size
Add zoom functions to ribbon
Configurable font size
Automatically fill scouting info from subsheets
Optional shorthand expansion
Shortcut to extend to next column of same side

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
Switch URL endpoints to production
Check for extraneous styles
Test all functions on PC
Test all functions on Mac
Rubberduck everything
Run template through the decompiler
	
# Future ideas
Window arranger for multiple screens
Save all style customizations to settings to survive updates
