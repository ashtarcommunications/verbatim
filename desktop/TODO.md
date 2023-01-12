# Bugs & New Features
Verify all document styles setup
Rebuild default keyboard shortcuts
Convert all library references to late binding
Finish caselist upload functions
Reorganize settings form
Normalize form styles
Convert all MacScript calls to AppleScriptTask
Standalone install checker tool
Plugin system for GetFromCitemaker, timer, OCR
Bluetooth add-on disabler
VTub refresh
Search functions
Mac OCR option (probably tesseract)
Rewrite update check
Add default event selection to settings and update stats form to pull from it
Troubleshooting form
Fix new speech dropdown
Rebuild ribbon
Add new macros/features to ribbon
Clean up Ribbon XML with proper IDs, keyboard shortcuts, descriptions
Rename CiteMaker functions to Cite Creator
Check all settings namespaces/names throughout
Build UI for autotext functions
Consider paragraph spacing adjustments
Remove unused styles macro
Search plugin
Mac VTub icons
Fix mac forecolor buttons on all forms
See if possible to check for internet before doing tabroom fetches or just better handle while disconnected
Figure out what's triggering Normal.dotm prompts
Tutorial Replacement
Make sure unighighlight with exception works with None setting
Replace vbNullString usage
Macro to auto switch cites from month/day to year for older files
Incorporate status bar progress meter from mac for tabroom-related functions
First line of VSCVerbatimSampleCard is test not text
Add ermo macro to remove underlining which is NOT highlighted.
Add ermo macro to shrink all text in separate paragraphs when card not condensed
Add ermo macro to reverse pilcrows to paragraphs
Add option to quickly toggle between retain paragraphs/pilcrows setting during condense macro
Dedicated hotkeys for paragraph integrity/pilcrows
Stats form frequently freezes and doesn't compute, especially on Mac
Add macro for Phillips to remove returns on paste (and figure out Mac version)
Try fixing the Condense macro when it doesn't find the first para break, because the search direction is set to "Up" - try changing it to down
Fix tilde adding comment in read mode with Teja Luburu modification to add Text:="Marked" and remove ActivePane line
Send selection/heading to end of file instead of cursor
Remap or remove keyboard shortcut for Remove All Emphasis (Ctrl-Shift-F10)
Fix F6 shortcut for Block on Mac (currently overridden by "Next Pane")
Change mac keyboard movement shorcuts to Cmd+Option+arrow
Adapt RepairCardFormatting to auto select a card or process all cards in file (plus test it more extensively)
Consider a native analytics style and a native remove analytics function
Consider a native undertags style – card notes that go under the tag
Refactor out Styles page on settings
Convert all formatting to built-in styles macro
Unshrink all macro
Judge view mode macro (remove emphasis and shrinking)

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

# Feature requests

Invisibility mode that works for large documents – even if that comes at the cost of deleting the unhighlighted, uncited, and untagged text, instead of hiding it (Teja’s code for this is below).
	A button to perform this modified invisibility mode for just the selected text, or just the paragraph that contains the cursor.
    These invisibility modes should delete card notes/undertags but NOT analytics. 

A button that saves two copies of the currently open file to the desktop – one in invisibility mode (and, if possible, a word count at the top) called “[name of document] Read”, and one with analytics and card notes/undertags removed called “[name of document] Send”.

Additional shrink text options:
    Function to shrink selected paragraphs that do not contain underlining or emphasis. This is to make paragraphs that are skipped completely take up less space on the screen.
    Function to shrink text that is not the normal size. This is because when I make [CHART/TABLE/FIGURE/ETC OMITTED] notations, I want to keep them unformatted 11 pt, but shrink the text around them. Currently I emphasize them, shrink the text, and then change them back to normal formatting by hand.

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
