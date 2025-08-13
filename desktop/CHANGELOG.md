# Verbatim Desktop Changelog
All notable changes to this project will be documented in this file. Changelogs for versions prior to 6.0.0 are hosted in a different repository.
<br /><br />
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [6.0.0-mini] - 2025-08-13

### Added
* Added a "Verbatim Mini" version stripping out all network access code and other integrations likely to trip antivirus scanners

## [6.0.0] - 2023-05-01

### Added
* Included a CONTRIBUTIONS.md detailing how to contribute to the project
* Allow hiding individual ribbon sections in the settings
* Macro to emphasize first letter of each word of selection, like "United States"
* Ability to set an exception color while standardizing highlighting
* Macro to remove non-highlighted underlined text
* Option to send to end of speech document instead of cursor
* Option to condense automatically on paste
* Options to override Pilcrows settings temporarily on condense
* Macro to unshrink all cards in the document
* Shortcut to select current heading and content
* Ribbon toggles and keyboard shortcuts for changing paragraph integrity and pilcrows settings
* Macro to normalize formatting across spaces/punctuation
* Quick Cards feature with user interface for creating/deleting/inserting small blocks of content
* Macros to convert all formatting to built-in styles and remove extraneous styles
* Macro to reformat existing cites to switch from year to month/day for older cites
* New OCR integrations, including a Mac option
* Macro to convert custom analytics styles into tags
* Macro to move current heading to the bottom of the current document
* Plugin system for overriding the built-in timer, OCR, or search
* New caselist upload feature to work with new openCaselist
* Integration with share.tabroom.com to allow privacy-first document sharing
* Everything Search plugin for better document searching integration
* Mac ribbon search integration with Spotlight
* New streamlined cross-platform tutorial
* Verbatim Flow template for flowing in Excel
* Verbatim Flow integration to send blocks to the flow at either current cell or column

### Changed
* Combined Mac & PC code bases
* Updated window arranger to work on Mac and with different dock positions
* Reordered ribbon to put important functions on the left to help with ribbon sections collapsing on small screen
* Updated ribbon icons to work on both Mac and PC, with custom icons where necessary
* Restyled VBA userforms to a more modern look and feel
* Rebuilt the VTub to be cross-platform
* Update check is now semver compatible
* Streamlined setup wizard on first start
* Shrink function now automatically handles multi-paragraph cards and includes options to ignore table/chart omissions
* Modified keyboard shortcuts on Mac to consistently use Command instead of Ctrl, fixed some bad defaults, and included alternates for broken F-keys
* Added choosing a default event to set e.g. default speech times
* Converted all library references to late binding for better backwards compatability
* Rewrote update check to work with new update server and not download new versions automatically to avoid tripping virus scanners
* Integrated the new speech dropdown with openCaselist for easier speech creation
* Reorganized settings form and added lots of new settings
* The tilde key now marks a card whenever you're in the active speech document, instead of relying on reading mode
* Option to unset the current speech document, so a speech doc can be used as a regular document
* Simplified troubleshooting form and moved some checks to the setup tool
* Unified selection modes for most core formatting macros (current card if no selection, selection if selected, or whole doc if at top)
* Modified most macros using the Find dialog to use a range and avoid changing users selection
* Cleaned up all the code to pass Rubberduck VBA code inspections

### Removed
* Removed Email feature, superceded by Tabroom sharing functionality
* Removed backfile converter because nobody uses old formats anymore
* Removed deprecated functions from the Ribbon
* Removed custom mouse icons from PC userforms because they're not compatible on Mac
* Removed all Win32 API declarations, now only support 64-bit
* Removed old PaDS functions since the service is retired
* Removed integrations with the old caselist

### Fixed
* Updated the WPM chart in the settings to reflect current speed averages
* Insert Header macro now pulls the correct names from the settings
* Bug in stripping "Speech" with auto save feature when file named Speech.docx
* Bug with duplicate path separators on auto saving
* Bug with speech doc names at 12PM or 12AM

### Security
* Rebuilt the installers to not disable macro security by default
* Removed troubleshooting functions that disable macro security or modify the registry
