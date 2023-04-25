---
sidebar_position: 2
id: formatting
title: Formatting Functions
---

Paste Text (default F2)
This macro will paste text from the clipboard as “unformatted” text. This should almost always be used rather than Ctrl+V for pasting in card text. A failure to use unformatted text will quickly add extraneous headings and styles, making the Navigation Pane unusable and slowing down your document.

Condense (default F3)
This macro will remove white space from the current selection, while optionally retaining paragraph integrity. By default, the paragraph integrity settings will replace each hard return with a small “pilcrow” paragraph sign (¶). This results in a single block of text, with small pilcrows scattered throughout at the original paragraph breaks. Alternately, it is possible to retain paragraph integrity while not using the pilcrows – the macro will then only eliminate extraneous white space. If “retain paragraph integrity” is turned off entirely in the settings, it will just condense the text to a single paragraph.

IMPORTANT NOTE: When cutting a PDF or similarly formatted document which includes line breaks after each line of text in a single paragraph, retaining paragraph integrity will result in too many pilcrows being inserted. The solution is to temporarily turn off the “retain paragraph integrity” setting while cutting that article, then turn it back on.

Cite (default F8)
The cite style is designed to be applied only to the last name and date – unlike the “tag” style it only applies to a single word or set of characters, not to the whole line.

Underlining (default F9)
The underline function is fairly self-explanatory. You can configure whether to bold underlined text as well in the Verbatim settings. The underline macro is also written to “toggle” between underlined and un-underlined text. This makes it easy to quickly correct underlining mistakes on the fly.

There is also an included “auto-underliner” on the ribbon – when turned on, this will immediately toggle the underlining for any highlighted text, without needing to press an additional shortcut key.

Emphasis (default F10)
By default, Emphasis will add a box around the current selection. Whether to use a box or just leave text bold (or larger) can be configured in the Verbatim settings.

Highlight (default F11)
Highlight will toggle the highlighting of the current selection on and off using the default highlighting color. The default color can be set with the “highlight color picker” on the ribbon. It is strongly recommended that you not use “light gray” as the highlight color. There is a known bug in Word which sometimes “loses” highlighting in saved files when this color is used.

Clear Formatting (default F12)
The Clear Formatting function will completely remove any formatting from the selection and return it to Normal text. The only thing it doesn’t remove is highlighting – this can be removed separately by toggling it with the highlighting function. When facing an intractable formatting problem, it is usually quickest to just clear the formatting of the offending text and start over.

Card Formatting Example

When correctly formatted, a card should look like the following:


Shrink Font (default Ctrl+8)
Reduces un-underlined parts of the current paragraph by progressively smaller font sizes, until it cycles back to the normal font size. Note that there must be at least some underlining in the paragraph to shrink the text.

Update Styles (default Ctrl+F2)
This will attempt to reformat the current document in your currently configured Verbatim template styles. Is mostly useful when opening a backfile that appears incorrectly, or after pasting in a card from a different source.

Select Similar Formatting (default Ctrl+F3)
Will select all portions of the document with formatting similar to the current selection – for example will select all “Tags” in the document so you can apply a uniform style change. Useful for quickly reformatting large sections of the file. Discussed in more detail below in the Converting Backfiles section.

Shrink Pilcrows
If you accidentally underline a pilcrow sign in your card, it will appear much larger and more annoying than it should. This macro will re-shrink and un-underline all pilcrows in the current paragraph. If run with the cursor at the very beginning of the document, it will shrink all pilcrows in the entire document.

Remove Blanks
Will delete any improperly formatted blank lines (accidentally formatted as a Heading Level), removing them from the Navigation Pane.

Remove Hyperlinks
Removes formatting from all hyperlinks in the document, to avoid inadvertently clicking on one.

Standardize Highlighting
Makes all highlighting in the current document the default color.

Insert Header
Creates a header based on the “Team Name” and “User Name” provided in the Verbatim settings.

Copy Previous Cite (default Ctrl+F8)
Will paste a copy of the previous cite at the current insertion point. Useful for cutting long documents with a lot of cards by the same author. Only works if the cite is contained in one paragraph (i.e. the author and date are not contained on a totally separate line). Works by finding the previous “Cite” style, so won’t work if you format your cites incorrectly.
