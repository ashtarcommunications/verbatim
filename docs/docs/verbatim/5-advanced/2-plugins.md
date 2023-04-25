---
sidebar_position: 2
id: plugins
title: Plugins
---

# PC

The easiest way to install all the plugins on a PC is to use the automated installer. That will install all of the following plugins:
* Verbatim Timer
* Capture2Text OCR
* Everything Search
* NavPaneCycle tool
* Get From Cite Creator tool

To manually install:

## Timer
Run the automated installer package, or place VerbatimTimer.exe in:
C:\Program Files\Verbatim\Plugins\VerbatimTimer.exe

You can choose to use any other external Timer program in place of the Verbatim Timer by configuring a path to a separate executable in the Verbatim settings.

## OCR
Install Capture2Text to one of the following two locations:
C:\Program Files\Verbatim\Plugins\OCR\
or
C:\Program Files\Capture2Text\

You also need to have the Window Snipping Tool installed (installed by default on most systems)

You can choose to use a differnet external OCR program instead of Capture2Text by configuring a path to a separate executable in the Verbatim settings. Note that when using an external OCR program, pressing the OCR ribbon button will launch that program, but won't be able to automatically paste the OCR'd result into your document.

## Everything Search
Install Capture2Text to one of the following two locations:
C:\Program Files\Verbatim\Plugins\Everything\
or
C:\Program Files\Everything\

You can choose to use a differnet external search program instead of Everything Search by configuring a path to a separate executable in the Verbatim settings. Note that when using an external saerch program, pressing the button for additional search results will only launch your external program, it won't be able to pass on the search terms to the program automatically.

## NavPaneCycle

Word’s Navigation Pane is very powerful, but can sometimes be very cluttered, especially in long files making full use of all four Verbatim heading levels. Unfortunately, the NavPane cannot be automated from within Word, and cannot be set to open “collapsed” to only Heading 1.

The first alternative is to manually change which Heading Levels are displayed in the Navigation Pane by right clicking anywhere in it and selecting “Show Heading Levels – Show Heading X.”

The downside of this approach is that it’s slow and repetitive.

To cover this gap until an official Microsoft solution, I’ve written a standalone program that enables a hotkey (Ctrl-`) to automatically cycle Headings 1-3 in the Nav Pane.  NavePaneCycle.exe is included as part of the paperless debate package download, or as a standalone file from the website.

The program doesn’t require any separate installation – it just needs to be called NavPaneCycle.exe and present in the Word Templates folder.

To use, you can either use the built-in shortcut key, Ctrl+` or press the NavPaneCycle button on the ribbon:  

The macro will take about a half second each time you press the shortcut (and you have to release both keys first). It will only work when Word is the active window, and when the Nav Pane is open.  Otherwise, it will do nothing.

IMPORTANT NOTE: It’s possible that NavPaneCycle will not work on your computer – if you find that it consistently doesn’t click in the right place to cycle the Nav Pane, or accidentally “demotes” sections of your file instead, then you’re probably out of luck.

You can also set an option in the Verbatim settings which will run NavPaneCycle automatically every time you open a new file, condensing the Nav Pane to only show Heading 1. This is somewhere between very convenient and very annoying, depending on how you look at it. Note that this sometimes conflicts with using the Virtual Tub – it’s not recommended you use both simultaneously.


## Others
Put the executables in the Verbatim Plugins directory:
C:\Program Files\Verbatim\Plugins
so
NavPaneCycle
C:\Program Files\Verbatim\Plugins\NavPaneCycle.exe

GetFromCiteCreator
C:\Program Files\Verbatim\Plugins\GetFromCiteCreator.exe

# Mac

On the Mac, plugins must be installed manually one by one.

## Timer
Download the VerbatimTimer.dmg package and drag VerbatimTimer.app to your Applications folder

## OCR
Install Tesseract using MacPorts (recommended) or Homebrew. For example, with MacPorts:

`ports install tesseract`

It is assumed tesseract will be installed on your path in one of the following locations:
/usr/bin/tesseract

## Search
Not supported on Mac. Instead, the built-in search uses the MacOS Spotlight tool.

## Others
NavPaneCycle is not supported on the Mac.
GetFromCiteCreator is built-in using Applescript, no separate installation is necessary
