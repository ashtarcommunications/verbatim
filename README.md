# Verbatim

## Overview

Verbatim is a free, open-source (GPL3) platform for paperless debate, with a variety of supporting tools and plugins. The main version is built as a template for Microsoft Word. It's primarily designed for usage in US-based high school and collegiate policy debate, but is usable by many other debate formats including Lincoln-Douglas, Public Forum, etc.

<br />

For usage documentation, see (https://paperlessdebate.com)

<br />

Desktop versions of Verbatim require a “full” version of Microsoft Office (Office 365 or desktop), which includes support for VBA macros. It will not work with Office Online, the Office Starter Pack, Office Home and Student, Office RT (e.g. for the Surface tablet), Office for iPad, Office for Android, or the version of Office in the Microsoft App Store.

## Project Structure

This repo contains a number of related tools, many of which need to be built and distributed separately. See the README in each directory for detailed information on each major subproject.

<br />

```
/desktop                     -- Verbatim for desktop Word, and associated tools/plugins
/docs                        -- Documentation site at docs.paperlessdebate.com, built with Docusaurus
/gdocs                       -- Google docs port of Verbatim
/timer                       -- Cross-platform debate timer built with Tauri
/owa                         -- Office Web Apps port of Verbatim
```

## Contribution Guidelines

See [CONTRIBUTING.md](CONTRIBUTING.md)
