#!/bin/bash
set -e

# Delete prior builds
rm -fr dist/*

mkdir -p dist/chrome
mkdir -p dist/firefox

# Copy the source tree with the browser-specific manifest file, and delete test files
cp -r src/* dist/chrome/
cp chrome/manifest.json dist/chrome/
rm dist/chrome/*.test.js
rm dist/chrome/test.*

cp -r src/* dist/firefox/
cp firefox/manifest.json dist/firefox/
rm dist/firefox/*.test.js
rm dist/firefox/test.*

# Package each browser build as a zip file
(cd dist/chrome; zip -r ../cite-creator-chrome.zip *)
(cd dist/firefox; zip -r ../cite-creator-firefox.zip *)
