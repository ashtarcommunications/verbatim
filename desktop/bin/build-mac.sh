#!/bin/sh

cp ../Debate.dotm ../install/mac/bundle/Verbatim.app/Contents/Resources/Debate.dotm
cp ../DebateStartup.dotm ../install/mac/bundle/Verbatim.app/Contents/Resources/DebateStartup.dotm
cp ../Verbatim.scpt ../install/mac/bundle/Verbatim.app/Contents/Resources/Verbatim.scpt
cp ../flow/Debate.xltm ../install/mac/bundle/Verbatim.app/Contents/Resources/Debate.xltm
cp ../CHANGELOG.md ../install/mac/bundle/Verbatim.app/Contents/Resources/CHANGELOG.md
touch ../install/mac/bundle/Verbatim.app
chmod -R 777 ../install/mac/bundle/Verbatim.app/Contents/Resources
chmod -R 777 ../install/mac/bundle/Verbatim.app/Contents/MacOS

pkgbuild --root "../install/mac/bundle" \
	--component-plist "../install/mac/Verbatim.plist" \
	--identifier "com.paperlessdebate.Verbatim" \
	--version "6.0.0" \
	--install-location "/Applications" \
	--scripts "../install/mac/scripts" \
	--ownership preserve \
	"../install/mac/Verbatim6.pkg"
