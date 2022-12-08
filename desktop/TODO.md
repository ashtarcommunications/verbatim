# New Stuff
Reorganize settings form
Convert all library references to late binding
All bug fixes from bug list
Normalize form styles
Convert all MacScript calls to AppleScriptTask
Mac installer check for Mac OS keyboard shortcuts
Standalone install checker tool
Plugin system for GetFromCitemaker, timer, OCR
Bluetooth add-on disabler
VTub refresh
Splash screen for first start
Search functions
Replacement for tutorial
Rewrite update check
Class documentation
Contributions.MD
Add default event selection to settings
Troubleshooting form
Fix new speech dropdown
Remove custom pointer ico's from form
Rebuild ribbon
Rename CiteMaker functions to Cite Creator
Check all settings namespaces/names throughout
Build UI for autotext functions
Consider paragraph spacing adjustments
Add getVisible callbacks
Search plugin
Mac VTub icons
Fix mac forecolor buttons on all forms
See if possible to check for internet before doing tabroom fetches

# Old Bugs
Caselist upload is broken with 401s
Update Mac Verbatim verbatimize to use a dynamic HD name, since "Macintosh HD" in the path isn't standard for all users
Spaces between some WPM in WPM Chart, others none
Accounts tab of Setup Wizard is misnamed lblStep3
Macro to box first letter of each word of highlight phrase, like "United States"
Macro to auto switch cites from month/day to year for older files
Insert Header macro pulls wrong names from settings - should be SchoolName/Name
Incorporate status bar progress meter from mac for tabroom-related functions
First line of VSCVerbatimSampleCard is test not text
For PC version, add Len(ActiveDocument.Name) > 11 to StripSpeech checks, to avoid bonking on Speech.docx, and add a trim to result
Paperless.NewSpeech doubles \ in the name, crashes autosave
Change update check to deal with multiple periods
Updating overrides style preferences, since those aren't saved in registry
Add an option to disable IE login, just in case it's bonking
Mac version troubleshooter lists PC path in the install warning
Considering modifying PaDS setup steps to add site to Local Intranet instead of just Trusted Sites.
Add ermo macro to remove underlining which is NOT highlighted.
Add ermo macro to shrink all text in separate paragraphs when card not condensed
Add ermo macro to reverse pilcrows to paragraphs
Add option to quickly toggle between retain paragraphs during condense macro

Tutorial button on Help form is disabled, even in 2011
Attach Template fails with an Invalid Command on 2016 when using the Verbatimize button - looks like the AppleScriptTask call is failing - maybe a bad path or a bad call to AttachTemplate from Toolbar.
Installer doesn't deal well with dual installs of 2011/2016
Stats form frequently freezes and doesn't compute, especially on Mac
Wiki upload crashing on some cites (Rajiv Movva)

Add macro for Phillips to remove returns on paste

Try fixing the Condense macro when it doesn't find the first para break, because the search direction is set to "Up" - try changing it to down

Speech Doc Namer sets 12PM to 12AM isntead. Change code to:
If Len(FileName) = 3 Then
  If Hour(Now) = 0 Then h = "12AM"
  If Hour(Now) = 12 Then h = "12PM"
  If Hour(Now) > 12 Then h = Hour(Now) - 12 & "PM"
  If Hour(Now) < 12 Then h = Hour(Now) & "AM"
  FileName = FileName & " " & Month(Now) & "-" & Day(Now) & " " & h
End If

Add new caselists to options

Fix tilde adding comment in read mode with Teja Luburu modification to add Text:="Marked" and remove ActivePane line

Wikification doesn't handle ñ characters correctly





# Flow TODO
send selected or heading to excel in cell or column
Send outline tag/cite from word to excel
send excel cell/column to word at cursor or end of doc
insert row(s) above/below current based on hotkey
insert cells above/below in same column
delete row
delete empty sheets
delete single sheet
easy move between sheets
add sheets quickly (aff/neg/etc)
toggle enter/alt-enter mode
highlight argument
Create flow for tabroom round/Speech from Verbatim
Label flows based on title in A1
Info tracking/scouting info
auto-move columns to chosen speech
macro to designate evidence for an argument
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


# Truf requests

A native analytics style and a native remove analytics function.

A native ‘undertags’ – card notes that go under the tag – style

Invisibility mode that works for large documents – even if that comes at the cost of deleting the unhighlighted, uncited, and untagged text, instead of hiding it (Teja’s code for this is below).
	A button to perform this modified invisibility mode for just the selected text, or just the paragraph that contains the cursor.
    These invisibility modes should delete card notes/undertags but NOT analytics. 

A button that saves two copies of the currently open file to the desktop – one in invisibility mode (and, if possible, a word count at the top) called “[name of document] Read”, and one with analytics and card notes/undertags removed called “[name of document] Send”.

A GUI for the autocorrect trick debaters sometimes do – e.g. where you type “econdefense” and that autocorrects to Walt 20. Current Word makes it a pain to keep this current. It’s usually faster to use this method than the Virtual Tub method, especially when the Virtual Tub gets large.  

Dedicated hotkeys for ‘condense with paragraph integrity’ and ‘condense without paragraph integrity.’

paste from a PDF that preserves paragraph integrity without having to copy paste paragraph-by-paragraph and without exporting to a word document… that would solve the problem and you would be my savior, but I imagine that is a stretch. 

Hotkey to apply the emphasis style to the first letter of each word in the selection – for acronyms (Sarah Lim’s code for this is below).

A form of highlighting that is easy to modify and yet is exempt from the ‘standardize highlighting’ function. Many people use different colors to recut cards and standardizing highlighting on accident can make it impossible to restore the original vs new highlighting.
	There is some ‘background color’ setting that MSU used to use, but their reason was to make it annoying to steal their cards, so the difficulty of making changes was a feature, not a bug. 

Hotkey to ‘repair underlining and highlighting’. This would undo a byproduct of cutting cards with my keyboard, which is that shift control arrow keys stops the selection at the end of words, leaving gaps between the formatting styles at the spaces. The fix would be:
	Look for space characters surrounded by the underlining style and underline them.
	Look for space characters with underlining on one side and emphasis on the other, and underline them.
	Look for space characters surrounded by the emphasis style and emphasize them.
	Look for space characters that are emphasized, with emphasis on one side and underlining on the other, and change them to underline
	Look for space characters that are emphasized, with emphasis on one side and neither emphasis nor underlining on the other, and clear their formatting
	Look for space characters that are underlined, with underlining or emphasis on one side and neither emphasis nor underlining on the other, and clear their formatting
	Look for space characters surrounded by highlighting and highlight them.
	Similar functions for punctuation---format comma space format, format period space format, format colon space format, format semicolon space format.
	I want to fix the gaps between formatting styles that result in spaces getting shrunk because they are not formatted.

Send card/header/hat/block to the end of the file. The bottom of the file is where I maintain my ‘cutting’ / ‘leftovers’ section.
		
Cutting cards with the keyboard on MacOS. The keyboard modifier to make the cursor move a word at a time on the mac is option. But option arrow keys is currently bound to DeleteHeading, MoveDown, MoveUp, and SendToSpeech. I have changed these to Command + Option + arrow keys which I think is a more sensible default given that these functions – as far as I know – are rarely used, and Command + Option is almost identically difficult to press.

Fix block style on Mac (F6). The F6 hotkey for the style does not work because F6 is bound to something else called “Next Pane.” I was not able to rebind “Next Pane” to something else and instead have been using Command F5 for the Block style.

Remove all emphasis (by default, Control Shift F10) should be bound to something else or unbound. While cutting cards with the keyboard and using Shift and Control to make selections, I frequently press Control Shift F10 on accident. Personally, I have unbound it.

Organize documents options for multiple screens.

Additional shrink text options:
    All shrink text options should go past 4 pt font.
    Function to shrink selected paragraphs that do not contain underlining or emphasis. This is to make paragraphs that are skipped completely take up less space on the screen.
    Function to shrink text that is not the normal size. This is because when I make [CHART/TABLE/FIGURE/ETC OMITTED] notations, I want to keep them unformatted 11 pt, but shrink the text around them. Currently I emphasize them, shrink the text, and then change them back to normal formatting by hand.
	
Hotkey to remove unused non-verbatim styles from the document.
            
Flowing
	Some information tracking stuff on the home page.
	A dropdown menu and button to create some number of case and off-case pages. These pages come with pre-set speech columns and the text is colored correctly.
	A button to name each tab based on the contents of A1.
	A drop-down and button to move the cursor to the top of the column corresponding to each speech.
	A macro to add x number of rows where the cursor is.
	A macro to highlight a selected cell (which I don’t think works on Mac because it conflicts with the hide app shortcut… haven’t bothered checking or fixing bc I don’t use it anyway).
	A macro to skip to the next sheet. 
        
	Currently when you have stuff in the clipboard and run the add rows macro, it goes haywire and also breaks the history for the undo button. I think this is because the macro is doing something weird with paste. Maybe storing, clearing, and restoring the clipboard would fix?
        
	A button to delete empty user-created sheets. For when the 1N overpromises on off-case arguments.
    A macro to designate an argument as having evidence attached. Could just be by adding a thick border to the bottom of the cell.  
        
	I want to be able to move stuff between the flow and speech doc, much like you currently can move stuff from a file to a designated speech doc. I have implemented these functions in AutoHotkey but it's janky as hell. This is so you can populate your flow with stuff from a doc, or populate your speech doc with stuff from your flow to avoid having to side-by-side it.  Specifically:
		Send tag + cite to flow
		Send selection to flow
		Send selection of flow to speech doc
		Send selected speech on flow to speech doc
			For sending stuff from the flow to the speech doc, I was able to make it so the opponents’ arguments appear as Level 3 headers, and your responses appear in the Analytics style as Level 4 headers.
					
A button similar to the ‘organize documents’ button in word that split screens the speech doc and the flow. 

    
Customize what appears on the debate tab.
            
Sometimes the verbatim installer puts Debate.dotm in a folder that is not the templates folder.

Stats function time estimate should allow you to customize WPM. 2022 average is 270-290 wpm. At my fastest, I was at 315 wpm.
