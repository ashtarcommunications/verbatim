# Verbatim README

## Things that do not work in Mac VBA
* CommandButton BackColor on UserForms (Use Forecolor as a replacement instead)
* #WIN64 compiler constant returns true on Mac (Have to use #If Mac Then <do nothing> Else <do PC only> #End If)
* Many ribbon icons
* Custom mouse pointers on forms
* GetSetting doesn't accept vbNullString for the default parameter, use "" instead
* Have to use replacements for most libraries, inc. XML, HTTP, Dictionary, ADODB, VBIDE
* PictureSizeMode on userform images doesn't resize the same on Mac as it does on PC

## Form UI Design

### Labels
BackColor &H00FFFFFF&
BackStyle 0 transparent
BorderColor &H00FFFFFF&
BorderStyle 0 None
Font Calibri 10
ForeColor &H00404040&
Special Effect 0 Flat

### Inputs
BackColor &H00FFFFFF&
BackStyle 1 opaque
BorderColor &H00A9A9A9&
BorderStyle 1 Single
Font Calibri 14
ForeColor &H80000008&
Height 24
Special Effect 0 Flat

### Buttons
BackColor &H00795C40&
BackStyle 1 Opaque
ForeColor &H00FFFFFF&
Height 30
Font Calibri 14

## Features I won't add
Shrink font below 4pt
Invisibility modes that actually delete content
Undertags style
Analytics style or anything that removes analytics automatically

## Notes For Docs
For Docs: Cursor selection philosophy, plead for 1-paragraph cite format