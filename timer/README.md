# Verbatim Timer

The Verbatim Timer is a cross-platform timer for debate events, meant to be integrated with the Verbatim template from paperlessdebate.com.

It's built as a Tauri app, which allows for a cross-platform native application without the overhead of bundling something like Electron.

It shares no code with the Debate Synergy timer by Alex Gulakov (GPL3), but it does take some design inspiration.

## Prerequisites
Install rust and dependencies

https://tauri.app/v1/guides/getting-started/prerequisites

Also requires Node.js

## Project Structure
/src-tauri -- includes the Tauri rust scaffolding and config files, as well as the application icons

/dist -- HTML/CSS/JS source for the actual application, all the business logic lives in `timer.js`

## Development
First, install npm dependencies:

`npm install`

Tauri is expecting a locally running dev server on port 8080 hosting the bundle. It's easiest to run webpack in dev mode (you can do some limited testing in the browser if desired):

`npm run dev`

Then run Tauri in dev mode to get the app running with hot reload:

`npm run tauri dev`

## Build
The final installation packages can be built with:

npm run tauri build

Cross-compilation isn't possible, so you have to build the package on each operating system separately and then get the final package from the `src-tauri/target` directory

Make sure to increase the version number in the config file and the version string in index.html when bulding a new version of the app.
