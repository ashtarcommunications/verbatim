;AHK to cycle Word 2010/2013 Nav Pane heading levels
;Not persistent - designed to be called from Word VBA
;Only cycles through first 3 headings, by design.
;Version 3.00
;June 2014 by Aaron Hardy

;Settings
#SingleInstance force
DetectHiddenText Off

;Wait half a second - lets most Word docs load before running
;Will still miss large documents and multiple docs
Sleep, 500

;Make sure Word is open with Doc Map displayed
IfWinActive, ahk_class OpusApp, MsoDockLeft
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

;Check if blue File Menu in Word 2013
PixelGetColor, Word2013, 6, 36

if Word2013 = 0x9A572B
{
  ;Right click on nav pane, move mouse and wait until next menu appears.  
  ;Increase sleep value from 250 if not working - this avoids lag bugs
  Sleep, 250
  Click right 100, 300
  Sleep, 250
  MouseMove, 190, 260
  Sleep, 350
  MouseMove, 60, 10

  ;Check for aqua checkmark
  PixelGetColor, H1Color, 11, 42
  PixelGetColor, H2Color, 11, 65
  
  ;If Heading 1 is checked, go to Heading 2
  if H1Color = 0x90A468
  {
    ;Set to Heading 2
    MouseMove, 70, 57
    Sleep, 250
    Click 70, 57
  }

  ;If Heading 2 is checked, go to Heading 3
  else if H2Color = 0x90A468
  {
    ;Set to Heading 3
    MouseMove, 70, 79
    Sleep, 250
    Click 70, 79
  }

  ;If Heading 3 or higher is checked, go back to Heading 1
  else {
    ;Set to Heading 1
    MouseMove, 70, 35
    Sleep, 250
    Click 70, 35
  }
}

else {
  ;Right click on nav pane, move mouse and wait until next menu appears.  
  ;Increase sleep value from 250 if not working - this avoids lag bugs
  Sleep, 250
  Click right 105, 270
  Sleep, 250
  MouseMove, 190, 245
  Sleep, 250
  MouseMove, 60, 10

  ;Check for black checkmark
  PixelGetColor, H1Color, 12, 37
  PixelGetColor, H2Color, 12, 59
  ;PixelGetColor, H3Color, 12, 81

  ;If Heading 1 is checked, go to Heading 2
  if H1Color = 0x000000
  {
    ;Set to Heading 2
    MouseMove, 70, 57
    Sleep, 250
    Click 70, 57
  }

  ;If Heading 2 is checked, go to Heading 3
  else if H2Color = 0x000000
  {
    ;Set to Heading 3
    MouseMove, 70, 79
    Sleep, 250
    Click 70, 79
  }

  ;If Heading 3 or higher is checked, go back to Heading 1
  else {
    ;Set to Heading 1
    MouseMove, 70, 35
    Sleep, 250
    Click 70, 35
  }
}

;Return mouse, turn on input and reset coordinate mode
CoordMode, Mouse, Screen
MouseMove, xpos, ypos
BlockInput Off
CoordMode, Mouse, Relative
return
}
