;AHK to cycle Word Nav Pane heading levels
;Not persistent - designed to be called from Word VBA
;Only cycles through first 3 headings, by design.
;Version 4.0.0
;2023 Ashtar Communications

;Settings
#SingleInstance force
DetectHiddenText Off

;Wait half a second - lets most Word docs load before running
;Will still miss large documents and multiple docs
Sleep, 500

;Make sure Word is open with Doc Map displayed
IfWinActive, ahk_class OpusApp, Navigation
{

;Get current mouse position
CoordMode, Mouse, Screen
MouseGetPos, xpos, ypos
CoordMode, Mouse, Relative

;Wait until key is unpressed to avoid stuck key
KeyWait vkC0
KeyWait Control

;BlockInput (Requires Admin)
BlockInput On

;Right click on nav pane, move mouse and wait until next menu appears.  
;Increase sleep value from 250 if not working - this avoids lag bugs
Sleep, 250
Click right 45, 305
Sleep, 250
MouseMove, 200, 305
Sleep, 350
MouseMove, 60, 10

;Check for aqua checkmark
PixelGetColor, H1Color, 19, 46
PixelGetColor, H2Color, 19, 75

;If Heading 1 is checked, go to Heading 2
if H1Color = 0xA4CD98
{
;Set to Heading 2
MouseMove, 70, 70
Sleep, 250
Click 70, 70
}

;If Heading 2 is checked, go to Heading 3
else if H2Color = 0x5EA649
{
;Set to Heading 3
MouseMove, 70, 100
Sleep, 250
Click 70, 100
}

;If Heading 3 or higher is checked, go back to Heading 1
else {
;Set to Heading 1
MouseMove, 70, 45
Sleep, 250
Click 70, 45
}

;Return mouse, turn on input and reset coordinate mode
CoordMode, Mouse, Screen
MouseMove, xpos, ypos
BlockInput Off
CoordMode, Mouse, Relative
return
}
