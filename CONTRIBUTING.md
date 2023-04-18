# Contributing

Community contributions to Verbatim and associated tools are welcome. However, changes will NOT be accepted via pull requests to this repository. The build process for a cross-platform VBA project contained in a Word template with a number of additional packaged tools is complex, difficult to automate, and doesn't lend itself well to a sane pipeline.

<br /><br />

Instead, if you'd like to contribute to these projects, contact support@paperlessdebate.com in advance to discuss your proposed changes and how to get them integrated. ALL changes to this repo require advance approval by the maintainer.

<br /><br />

Please also note that not all suggestions or feature additions will be accepted if they aren't useful for the community at large, or if they introduce features which don't align with the projects goals. Verbatim has many mechanisms for individual customization and extension, remember that your preferences may not be the same as other users.

<br /><br />

This repo has a number of discrete projects, including the desktop version of Verbatim, the Verbatim timer, and various setup tools/installer, etc. To get an overview of the project, read the README.

<br /><br />

For desktop Verbatim, note that the final template is NOT built from the /src directory. The source of truth for the current template is Debate.dotm - the /src directory is an export of the code modules from the Word template, and exists only to make it easier to use a git diff to understand what's contained in each commit.

<br /><br />

VBA code should also pass all Rubberduck VBA (https://rubberduckvba.com/) code inspections using the rubberduck.xml settings in the desktop directory.
