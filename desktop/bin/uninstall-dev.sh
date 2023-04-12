#!/bin/sh

if [ -d ~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized" ]
then
	rm ~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/Debate.dotm"
	rm ~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/DebateStartup.dotm"
	rm ~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/Debate.xltm"
else
	rm ~/"Library/Group Containers/UBF8T346G9.Office/User Content/Templates/Debate.dotm"
	rm ~/"Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/DebateStartup.dotm"
	rm ~/"Library/Group Containers/UBF8T346G9.Office/User Content/Templates/Debate.xltm"
fi

rm ~/"Library/Application Scripts/com.microsoft.Word/Verbatim.scpt"
rm ~/"Library/Application Scripts/com.microsoft.Excel/Verbatim.scpt"
rm -fr /Applications/Verbatim.app
rm -fr /Applications/VerbatimTimer.app

