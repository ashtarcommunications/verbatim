# Verbatim Desktop

## Overview
This section of the repo is for Verbatim for desktop versions of Microsoft Word. The project is made up of a few discrete tools:
1) The Verbatim Word template, comprised of Debate.dotm (the core template), DebateStartup.dotm (for bootstrapping the main template), and Verbatim.scpt, for Applescript usage on the Mac
2) Verbatim Flow, an Excel template for flowing, in Debate.xltm
3) A Mac .app bundle wrapping the template, with a .pkg installer built by pkgbuild and a convenience uninstall .app bundle
4) A PC .exe installer wrapping the template, built with NSIS
5) A PC plugins installer (also NSIS), which bundles several plugins:<br />
    a) Verbatim Timer, a cross-platform debate timer also in this repo<br />
    b) Capture2Text for OCR<br />
    c) Everything Search, for search<br />
    d) GetFromCiteCreator - an AutoHotKey script compiled to exe that fetches cite info from the Cite Creator extension in Chrome<br />
    e) NavPaneCycle - an AutoHotKey script compiled to exe that cycles through the Word Navigation Pane levels<br />
6) A PC setup check tool built with Visual Studio that helps administer macro security settings
7) A Mac setup check tool built with Applescript that administers macro security and various other Mac system settings
8) A "Mini" version of the two Word templates with features most likely to be flagged by antivirus removed.

<br />

Note that the final template is NOT built from the /src directory. The source of truth for the current template is Debate.dotm - the /src directory is an export of the code modules from the Word template, and exists only to make it easier to use a git diff to understand what's contained in each commit.

<br />


## Project Structure
```
/assets                           -- Images, icons, libraries, etc. for building projects
+---/icons                        -- Application and ribbon icons, including scripts to generate ribbon-usable icons from Font Awesome
+---/nsis                         -- NSIS plugins needed for building PC installers
+---/tutorial                     -- Ribbon images for embedded Verbatim tutorial

/bin                              -- Scripts to install/build/release. Includes VBADecompiler tool.

/flow                             -- Verbatim Flow template for Microsoft Excel
+---/src                          -- Exported code modules from the Excel template
+---Debate.xltm                   -- Self-contained Excel template for flowing, source of truth for Verbatim Flow

/install                          -- Installer packages, must be build on each platform separately
+---/mac                          -- Mac installer
+---+---/bundle                   -- Contains .app bundle, must have new versions of files copies to it
+---+---/scripts                  -- postinstall script used by the Verbatim.pkg installer to distribute the .app and template files
+---+---/VerbatimUninstall.app    -- Convenience app bundle for uninstalling all Verbatim related files on Mac
+---+---/Verbatim.plist           -- plist to feed to pkgbuild to build the Verbatim.pkg installer
+---+---/Verbatim6.pkg            -- Final Mac installer
+---+---/VerbatimUninstall.zip    -- Distribution .zip for the Mac uninstaller .app
+---/pc                           -- PC installer
+---+---/Verbatim6.exe            -- Final PC installer
+---+---/Verbatim6.nsi            -- NSIS installer script to generate Verbatim6.exe

/mini	                          -- "Mini" version of the 2 main Word templates, currently created manually

/plugins                          -- Plugins for various Verbatim features.
+---/ocr                          -- Latest .zip distribution of Capture2Text
+---/search                       -- Latest executable for Everything Search
+---GetFromCiteCreator.ahk        -- AutoHotKey source for GetFromCiteCreator plugin
+---GetFromCiteCreator.exe        -- Compiled build for GetFromCiteCreator plugin
+---NavPaneCycle.ahk              -- AutoHotKey source for NavPaneCycle plugin
+---NavPaneCycle.exe              -- Compiled build for NavPaneCycle.exe
+---VerbatimPlugins.exe           -- Final PC installer for plugins
+---VerbatimPlugins.nsi           -- NSIS installer script to generate VerbatimPlugins.exe

/release                          -- Release archive, one subdirectory per semver release, which includes all finished distributables

/setup                            -- Verbatim Setup Check tools
+---Verbatim Setup Check          -- Visual Studio project for building PC setup check tool
+---VerbatimSetupCheck.app        -- Mac .app bundle for Verbatim Setup Check tool
+---VerbatimSetupCheck.zip        -- Distributable .zip for the Mac setup tool

/src                              -- Exported code modules for the Word template from Rubberduck VBA Code Explorer. For reference/git diffs only.

CHANGELOG.md                      -- Changelog for the desktop project and associated tools
Debate.dotm                       -- Main Verbatim template for desktop Word, source of truth for Verbatim VBA code
DebateStartup.dotm                -- Global Word Startup template for bootstrapping Debate.dotm in "Always On" mode
rubberduck.xml                    -- RubberDuck VBA code inspection definitions
Verbatim.scpt                     -- Applescript file for use with Mac Verbatim
```

## Build/Release Process
1) Ensure Debate.dotm compiles on both PC and Mac, and passes all Rubberduck VBA code inspections
2) Export current state of code modules to /src and /flow/src with Rubberduck VBA Code Explorer
3) On PC: `cd bin && build-pc.bat`, or manually run Debate.dotm, DebateStartup.dotm, and Debate.xltm through VBADecompiler.exe and delete backup artifacts
4) Build PC installer and PC Plugin Installer with NSIS (also included in `build-pc.bat`)
5) On Mac, run `build-mac.sh` or manually copy new versions of files to the Mac installer bundle, and run pkgbuild to built the .pkg Mac installer
6) Ensure all other tools are built, including Plugins, plugin installer, setup tools (PC and Mac), Mac uninstaller, and Timer
7) Run `bin/release.bat x.x.x` to copy all build artifacts to the `/release` tree
8) Manually create a "Mini" release - incorporate any template changes into the "Mini" versions then manually create a release folder renaming them to Debate.dotm and DebateStartup.dotm for distribution

## Things that do not work in Mac VBA
The Mac version of Word and the Mac VBA runtime have a lot of bugs, feature limitations, and unexpected behavior relative to the PC. This is a list of Mac-specific gotchas that have been worked around, and should be kept in mind when adding new features:
* #WIN64 compiler constant inexplicably returns true on Mac. Instead, use `#If Mac Then <do nothing> Else <do PC only> #End If`
* CommandButton .BackColor property on UserForms doesn't work, so on Mac use .Forecolor as a replacement instead
* Many ribbon icons are missing, so effort has been made to only use cross-platform available icons, or custom icons where necessary
* Custom mouse pointers on Userforms will cause compilation errors on Mac that fail silently, and should not be used
* GetSetting doesn't accept vbNullString for the default parameter, so have to use "" instead
* Mac does not have Windows-specific libraries, such as XML, HTTP, Dictionary, ADODB, VBIDE, etc. Instead, replacements have been built for all functionality, usually leveraging AppleScriptTask and shell scripts, sometimes using compatibility shims
* .PictureSizeMode on Userform images doesn't resize the same on Mac as it does on PC, so have to be careful with form element placement and sizing
* System.PrivateProfileString doesn't write to a separate ini file, it writes to the plist, so there's no easy way to import/export settings files, and this feature has been disabled on Mac
* AppleScriptTask has undocumented pipe buffer limits (~16K) on the length of the return value of a script invocation, necessitating workarounds when using curl with large payloads
* Some F-keys are completely stolen by MacOS (e.g. F6) and cannot be disabled, so alternate keyboard shortcuts have been included instead
* Application.OnKey in Excel can't use the Command key for shortcuts after Excel 2011, and many other key combos don't work, so alternate shortcuts have been included for Mac

## Features I won't add
Some features that are frequently requested will not be added to Verbatim. The short explanation is that these features are almost always used to encourage bad practices which are anti-competitive, anti-educational, and anti-accessibility. In general, this project supports the idea of an open exchange of information, and making your arguments as clear and accessible as possible. If your motivation for wanting a feature is to obfuscate your arguments or make them more difficult for the other team to understand, that feature won't be added. If you need these techniques to win, you should focus on getting better at debate, not finding ways to make debate worse. These include:
* A separate "Analytics" style, or any macros that remove analytics automatically
* Shrinking font below 4pt, which is already the limit on readability
* Invisibility modes that actually delete content, as opposed to just hiding it
* An "undertags" style for notes - this is due to creating complications with automatic card parsing when separating the tag and cite

## VBA Userform UI Design
All Userforms should follow a consistent style. It's easiest to consult an existing form, and copy elements from there. A few guidelines:

### Labels
* BackColor &H00FFFFFF&
* BackStyle 0 transparent
* BorderColor &H00FFFFFF&
* BorderStyle 0 None
* Font Calibri 10
* ForeColor &H00404040&
* Special Effect 0 Flat

### Inputs
* BackColor &H00FFFFFF&
* BackStyle 1 opaque
* BorderColor &H00A9A9A9&
* BorderStyle 1 Single
* Font Calibri 14
* ForeColor &H80000008&
* Height 24
* Special Effect 0 Flat

### Buttons
* BackColor &H00795C40&
* BackStyle 1 Opaque
* ForeColor &H00FFFFFF&
* Height 30
* Font Calibri 14

## Other libraries
Verbatim is bundled with a number of other libraries used for plugins and builds. These are included with their own licenses:
* VBADecompiler.exe
* Capture2Text
* EverythingSearch
