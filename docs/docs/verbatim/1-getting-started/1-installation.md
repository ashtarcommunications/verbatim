---
sidebar_position: 1
id: installation
title: Installation
---

# Installation

To download the latest version (full or mini), see [paperlessdebate.com](https://paperlessdebate.com)

Problems with installation? Check the [FAQ](../faq).

Getting a virus alert? See [Verbatim and Antivirus](./virus).

If you have issues with the automated installer, use the manual installation instructions below.

If you're having issues installing the full version, try Verbatim Mini, which is a stripped down version of Verbatim with all the core functionality, but with features removed (like Tabroom integrations and automatic updates) that are most likely to trip antivirus scanners. If you’re having issues with the full version or have particularly restrictive school IT, try this version.

## Automated Installation

The automatic installer for your operating system should work for the majority of people. Only the full version has an automated installer. If you need the Mini version, you must install it manually with the instructions below.

Then, make sure to read the section on [Security](./security) to understand how to adjust your macro security settings for use with Verbatim.

On a PC, you also likely want to download the Plugins installer to automatically install all Verbatim [plugins](../advanced/plugins) at once, including the Vebratim Timer, OCR, search and more.

On Mac, you may need to allow the installer to run as an unidentified developer. After running the installer, go to the apple menu, System Preferences - Privacy & Security, and grant an exception for the Verbatim installer.

For more detailed instructions, see:

https://support.apple.com/guide/mac-help/open-a-mac-app-from-an-unidentified-developer-mh40616/mac

After running the automated installer, you can open Verbatim with the shortcut on your Desktop (or Start Menu/Applications folder). Or, you can open Word and select "New from Template" to open a file based on the "Debate" template.

You may also run one of the Setup Check tools described in the section on [Security](./security) to ensure you've installed Verbatim correctly, and to help you manage your macro security settings.

## Manual Installation

If you’re using a school computer that disallows installing programs, or have other issues with the installer, you can use the instructions below to perform a manual installation. The procedure is the same for the full version and the Mini version, just make sure to download the correct template files from the downloads page.

Because Verbatim is a template for Word, to "install" it you just need to ensure the template files are placed in the correct folders. There are only 2 important files (3 on Mac):

Debate.dotm - this must be placed in your Office "Templates" folder.
DebateStartup.dotm - this must be placed in your Office "STARTUP" folder.
Verbatim.scpt (Mac only) - must be placed in your Application Scripts folder.

### PC

First, download the latest template files, `Debate.dotm` and `DebateStartup.dotm`, from the downloads page. Choose whether you want the full version, or the Mini version. Make sure that the files are saved as e.g. “Debate.dotm” and not “Debate (1).dotm” if you have accidentally downloaded more than one copy.

1) Move `Debate.dotm` to your Templates folder, usually located at:

`c:\Users\[Your Name]\AppData\Roaming\Microsoft\Templates\`

You can type `%APPDATA%` into the address bar of a Windows Explorer window to get most of the way there.

If you'd also like to install Verbatim Flow for Excel, put a copy of the `Debate.xltm` template in your Templates folder as well.

2) Move `DebateStartup.dotm` to your STARTUP folder, usually located at:

`c:\Users\[Your Name]\AppData\Roaming\Microsoft\Word\STARTUP\`

That’s it, Verbatim is installed! For ease of use, create a “shortcut” to Debate.dotm on your desktop (not a copy of the file itself).

### Mac

First, download the latest template files, `Debate.dotm` and `DebateStartup.dotm`, and the Verbatim script file `Verbatim.scpt` from the downloads page. Choose whether you want hte full version, or the Mini version. Make sure that the files are saved as e.g. “Debate.dotm” and not “Debate (1).dotm” if you have accidentally downloaded more than one copy.

Note that the following folders may be hidden by default, so you can use `Cmd+Shift+G` to go directly to the locations listed.

1) Move Debate.dotm to your Templates folder, usually located at:

`~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates/`

If you'd also like to install Verbatim Flow for Excel, put a copy of the `Debate.xltm` template in your Templates folder as well.

2) Move DebateStartup.dotm to your Startup folder, usually located at:

`~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/`

3) Move `Verbatim.scpt` to:

`~/Library/Application Scripts/com.microsoft.Word/`

You may need to manually create the `com.microsoft.Word` folder for Verbatim.scpt – be careful to use the exact punctuation and capitalization above.
