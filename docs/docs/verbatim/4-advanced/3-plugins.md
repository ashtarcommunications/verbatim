---
sidebar_position: 3
id: plugins
title: Plugins
---

Verbatim has a number of plugins which extend the functionality of the template. You can also manually override many of these plugins to use your own program for a Timer, OCR, search, etc.

Note that some plugins are PC-only, and will not work on the Mac.

## PC Installation

The easiest way to install all the plugins on a PC is to use the automated installer. That will install all of the following plugins:
* Verbatim Timer
* Capture2Text OCR
* Everything Search
* NavPaneCycle tool
* Get From Cite Creator tool

Note that the Timer is included as part of the Plugins pack, so you don't need to install the standalone timer as well.

If for some reason you need to manually install plugins:

### Timer Manual Installation
Place VerbatimTimer.exe in:

`C:\Program Files\Verbatim\Plugins\VerbatimTimer.exe`

You can choose to use any other external Timer program in place of the Verbatim Timer by configuring a path to a separate executable in the Verbatim settings.

### OCR Manual Installation
Install Capture2Text to one of the following two locations:

`C:\Program Files\Verbatim\Plugins\OCR\`

or

`C:\Program Files\Capture2Text\`

You also need to have the Window Snipping Tool installed (installed by default on most systems)

You can choose to use a differnet external OCR program instead of Capture2Text by configuring a path to a separate executable in the Verbatim settings. Note that when using an external OCR program, pressing the OCR ribbon button will launch that program, but won't be able to automatically paste the OCR'd result into your document.

### Everything Search Manual Installation
Install Everything Search to one of the following two locations:

`C:\Program Files\Verbatim\Plugins\Everything\`

or

`C:\Program Files\Everything\`

You can choose to use a differnet external search program instead of Everything Search by configuring a path to a separate executable in the Verbatim settings. Note that when using an external saerch program, pressing the button for additional search results will only launch your external program, it won't be able to pass on the search terms to the program automatically.

### Other Manual Installation
Put the executables for other plugins in the Verbatim Plugins directory:

`C:\Program Files\Verbatim\Plugins\NavPaneCycle.exe`

`C:\Program Files\Verbatim\Plugins\GetFromCiteCreator.exe`

## Mac Installation

On the Mac, plugins must be installed manually one by one.

### Timer
Download the `VerbatimTimer.pkg` package, which will isntall VerbatimTimer.app to your Applications folder.

### OCR
OCR on Mac requires installation of Tesseract, an open source OCR library. You can install Tesseract using either [MacPorts](https://www.macports.org) (recommended) or [Homebrew](https://brew.sh). For example, with MacPorts:

`ports install tesseract`

It is assumed tesseract will be installed on your path in one of the following locations:
* `/opt/local/bin/tesseract`
* `/usr/local/opt/tesseract`
* `/opt/homebrew/bin/tesseract`
* `usr/local/bin/tesseract`
* `/usr/bin/tesseract`

## Search
Not supported on Mac. Instead, the built-in search uses the MacOS Spotlight tool.

## Others
NavPaneCycle is not supported on the Mac.

GetFromCiteCreator is built-in using Applescript, no separate installation is necessary
