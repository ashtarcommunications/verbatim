---
sidebar_position: 3
id: selection
title: Selection Modes
---

Most formatting macros operate with a similar method of deciding what portion of the document to operate on.

Selection philosophy:
* Always respect user selection where possible. If text is selected, the macro will operate on that
* Otherwise, operate on the current heading - if in a card or on a tag, just operate on that card
* If on a larger heading (Pocket, Block, Hat), operate on that heading level
* If at the very beginning of the document, operate on the entire document