# Bugs & New Features
Verify all document styles setup
Rebuild default keyboard shortcuts
Convert all library references to late binding
Finish caselist upload functions
Reorganize settings form
Normalize form styles
Convert all MacScript calls to AppleScriptTask
Plugin system for GetFromCiteCreator, timer, OCR
VTub refresh
Search functions
Mac OCR option (probably tesseract)
Rewrite update check
Troubleshooting form
Fix new speech dropdown
Rebuild ribbon and add new macros/features/shortcuts
Clean up Ribbon XML with proper IDs, keyboard shortcuts, descriptions
Check all settings namespaces/names throughout
Build UI for autotext functions
Consider paragraph spacing adjustments
Remove unused styles macro
Search plugin
Mac VTub icons
See if possible to check for internet before doing tabroom fetches or just better handle while disconnected
Tutorial Replacement
Make sure unighighlight with exception works with None setting
Macro to auto switch cites from month/day to year for older files
Incorporate status bar progress meter from mac for tabroom-related functions
First line of VSCVerbatimSampleCard is test not text
Rewrite Find blocks to use a range where possible
Add option to quickly toggle between retain paragraphs/pilcrows setting during condense macro
Stats form frequently freezes and doesn't compute, especially on Mac
Try fixing the Condense macro when it doesn't find the first para break, because the search direction is set to "Up" - try changing it to down
Mark card feature for Mac
Fix F6 shortcut for Block on Mac (currently overridden by "Next Pane")
Adapt RepairCardFormatting to auto select a card or process all cards in file (plus test it more extensively)
Consider a native analytics style and a native remove analytics function
Consider a native undertags style – card notes that go under the tag
Replace all formatting with built-in styles macro
Judge view mode macro (remove emphasis and shrinking)
Unify card/doc selection modes (shrink, condense, uncondense, shrink pilcrows, cite-ify, convert dates, etc.)
Save "share version" of current document with no analytics/notes

# Flow template
send selected or heading to excel in cell or column
Send outline tag/cite from word to excel
send excel selection/cell/column to word at cursor or end of doc
insert row(s) above/below current based on hotkey (make sure it doesn't screw up clipboard/undo)
insert cells above/below in same column
delete row
delete empty sheets
delete single sheet
easy move between sheets
add sheets quickly (aff/neg/etc), both hot key and ribbon menu
toggle enter/alt-enter mode
highlight argument/cell
Create flow for tabroom round/Speech from Verbatim
Label flows based on title in A1 (optionally disable)
Info tracking/scouting info
auto-move cursor to columns of chosen speech on all sheets
macro to designate evidence for an argument/cell
auto-organize windows for flow/speech side-by-side
paste special or expand by crlf
vtub for analytics
Add/remove numbers for args automatically
Merge/group cells
Move up/down like with blocks
Configurable for different events
Keyboard shortcuts cheat sheet
Keep f2 shortcut for entering cell
change row/column size
add zoom functions to ribbon
configurable font size
automatically fill scouting info from subsheets
Optional shorthand expansion
shortcut to extend to next column of same side

# New Tools
PC setup tool
Mac setup tool
PC Installer
	Fix template getting put into admin templates folder when run with UAC
Mac Installer
	Mac installer check/disbale Mac OS keyboard shortcuts

# Pre-release QA
Double check all Mac-specific macro modifications made it into the combined version
Test all functions on PC
Test all functions on Mac
Rubberduck everything
Run template through the decompiler
	
# Future ideas
Window arranger for multiple screens
Save all style customizations to settings to survive updates
