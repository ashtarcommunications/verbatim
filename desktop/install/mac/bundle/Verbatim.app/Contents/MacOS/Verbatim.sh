#!/bin/sh

if [ -d ~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized" ]
then
	VERBATIM=~/"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/Debate.dotm"
else
	VERBATIM=~/"Library/Group Containers/UBF8T346G9.Office/User Content/Templates/Debate.dotm"
fi

osascript \
        -e 'set VerbatimPath to POSIX file "'"$VERBATIM"'" as string' \
        -e 'if application "Microsoft Word" is running then' \
                -e 'tell application "Microsoft Word"' \
                        -e 'activate' \
                        -e 'set myDoc to (create new document attached template VerbatimPath)' \
			-e 'run VB macro macro name "Startup.Start"' \
                -e 'end tell' \
        -e 'else' \
                -e 'tell application "Microsoft Word"' \
                        -e 'activate' \
                        -e 'set myDoc to (create new document attached template VerbatimPath)' \
			-e 'run VB macro macro name "Startup.Start"' \
                        -e 'delay 2' \
                        -e 'if exists document "Document1" then close document "Document1"' \
                -e 'end tell' \
        -e 'end if' \
> /dev/null 2>&1 &

exit
