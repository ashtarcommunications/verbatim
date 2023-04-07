#!/bin/sh

if [ -d ~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized" ]
then
	SCRIPTFILE=~/"Library/Application Scripts/com.microsoft.Word/Verbatim.scpt"
	VERBATIM=~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/Debate.dotm"
	STARTUP=~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/DebateStartup.dotm"
	FLOW=~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/Debate.xltm"
else
	SCRIPTFILE=~/"Library/Application Scripts/com.microsoft.Word/Verbatim.scpt"
	VERBATIM=~/"Library/Group Containers/UBF8T346G9.Office/User Content/Templates/Debate.dotm"
	STARTUP=~/"Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/DebateStartup.dotm"
	FLOW=~/"Library/Group Containers/UBF8T346G9.Office/User Content/Templates/Debate.xltm"
fi

RELEASEDATE="20230407"

DIR=$(cd $(dirname "$0"); pwd)
cd "$DIR"

if [ -e "$SCRIPTFILE" ]
then
	DATEMODIFIED=`stat -f "%Sm" -t "%Y%m%d" "$SCRIPTFILE"`
	if [ "$DATEMODIFIED" -le "$RELEASEDATE" ]
	then
		rm -f "$SCRIPTFILE"
		cp -p ../Resources/Verbatim.scpt "$SCRIPTFILE"
	fi
else
	if [ -d ~/"Library/Application Scripts/com.microsoft.Word" ]
	then
		cp -p ../Resources/Verbatim.scpt "$SCRIPTFILE"
	else
		mkdir -p ~/"Library/Application Scripts/com.microsoft.Word"
		cp -p ../Resources/Verbatim.scpt "$SCRIPTFILE"
	fi
fi

if [ -e "$VERBATIM" ]
then
	DATEMODIFIED2=`stat -f "%Sm" -t "%Y%m%d" "$VERBATIM"`
	if [ "$DATEMODIFIED2" -ge "$RELEASEDATE" ]
	then
		osascript -e 'set VerbatimPath to POSIX file "'"$VERBATIM"'" as string' -e 'tell application "Microsoft Word"' -e 'activate' -e 'set myDoc to (create new document attached template VerbatimPath)' -e 'run VB macro macro name "Startup.Start"' -e 'end tell' > /dev/null 2>&1 &
	else
		rm -f "$VERBATIM"
		rm -f "$STARTUP"
		cp -p ../Resources/Debate.dotm "$VERBATIM"
		cp -p ../Resources/DebateStartup.dotm "$STARTUP"
		osascript -e 'set VerbatimPath to POSIX file "'"$VERBATIM"'" as string' -e 'tell application "Microsoft Word"' -e 'activate' -e 'set myDoc to (create new document attached template VerbatimPath)' -e 'run VB macro macro name "Startup.Start"' -e 'end tell' > /dev/null 2>&1 &
	fi
	
else
	cp -p ../Resources/Debate.dotm "$VERBATIM"
	cp -p ../Resources/DebateStartup.dotm "$STARTUP"
	
	osascript -e 'set VerbatimPath to POSIX file "'"$VERBATIM"'" as string' -e 'tell application "Microsoft Word"' -e 'activate' -e 'set myDoc to (create new document attached template VerbatimPath)' -e 'run VB macro macro name "Startup.Start"' -e 'end tell' > /dev/null 2>&1 &
fi

if [ -e "$FLOW" ]
then
	DATEMODIFIED=`stat -f "%Sm" -t "%Y%m%d" "$FLOW"`
	if [ "$DATEMODIFIED" -le "$RELEASEDATE" ]
	then
		rm -f "$FLOW"
		cp -p ../Resources/Debate.xltm "$FLOW"
	fi
else
	cp -p ../Resources/Debate.xltm "$FLOW"
fi

exit
