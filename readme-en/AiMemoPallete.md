# AiMemoPallete

[![Direct](https://img.shields.io/badge/Direct%20Link-AiMemoPallete.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AiMemoPallete.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiMemoPallete.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A floating memo palette for Illustrator.

### Importing

- Collects text from the selected objects (Append by default, or Replace). Appending inserts one blank line between the existing text and the new text
- Text inside groups, clip groups and symbols is included, expanding nested levels
- With nothing selected, the clipboard is pasted temporarily (`pasteInAllArtboard`), the text is collected, and the pasted objects are deleted; per-artboard duplicates are removed
- Multiple texts are ordered by their position on the canvas, top to bottom (left to right at the same height). Effectively empty frames are ignored
- Voiced and semi-voiced marks are normalized to NFC on import
- Every button carries a tooltip (`helpTip`)

### Editing and copying

- Remove blank lines strips blank lines (including whitespace-only lines) from the text field
- Remove line breaks joins the whole text field into a single line
- Copy all copies the text field to the system clipboard (a temporary text frame plus `app.copy()`, executed through BridgeTalk)
- Buttons dim automatically when there is nothing to act on (Remove blank lines when there are none, Remove line breaks when there are none)

### Script info

- Version: v1.1.2
