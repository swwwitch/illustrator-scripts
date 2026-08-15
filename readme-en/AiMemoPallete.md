# AiMemoPallete

[![Direct](https://img.shields.io/badge/Direct%20Link-AiMemoPallete.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AiMemoPallete.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiMemoPallete.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A floating memo palette for Illustrator.

- Collects text from the selected objects or the clipboard and keeps it in one place.
- Tidies blank lines and line breaks, then saves the memo to a text file or copies it to the clipboard.
- Being a palette, it stays open while you keep working on the document. The content and window position are restored on the next launch.

### How to use

1. Run the script to open the floating palette.
2. Select objects that contain text and click Load. With nothing selected, the text is read from the clipboard.
3. Tidy the text with Remove Blanks / Remove Breaks; you can also type and edit directly.
4. Click Save to write a text file, or Copy All to put the memo on the clipboard.

### Buttons

| Item | Description |
| --- | --- |
| Mode | Append (default) adds the loaded text after one blank line. Replace swaps the whole field |
| Load | Loads text from the selected objects; with nothing selected, loads text from the clipboard |
| Remove Blanks | Removes blank lines (including whitespace-only lines) from the memo |
| Remove Breaks | Removes every line break, joining the memo into a single line |
| Save | Saves the memo as a UTF-8 text file (.txt) |
| Copy All | Copies the memo to the clipboard |
| Clear | Empties the memo and restarts the script |

Every button carries a tooltip (`helpTip`), and buttons dim automatically when there is nothing to act on (no text, no blank lines, no line breaks).

### Loading in detail

- Text inside groups, clip groups and symbols is included, expanding nested levels.
- Multiple texts are ordered by their position on the canvas, top to bottom (left to right at the same height).
- Effectively empty text frames are ignored.
- Decomposed voiced / semi-voiced marks are normalized to their composed form on import.
- With nothing selected, the clipboard is pasted into the document (`pasteInAllArtboard`), the text is read, and the pasted objects are removed. Per-artboard duplicates are dropped.

### Saving

- A footer with the source document name and the save timestamp is appended to the file.
- By default the file goes to the Desktop as `memo-<document name>-<yyyymmdd>.txt` (`SAVE_LOCATION_MODE`).
- Saving does not close the palette: the content and window position are preserved.

### Settings (top of the script)

| Constant | Default | Description |
| --- | --- | --- |
| `PALETTE_OPACITY` | `0.97` | Opacity of the palette |
| `SAVE_LOCATION_MODE` | `'desktop'` | `'desktop'` always saves to the Desktop; `'dialog'` asks for the location and name |
| `CLEAR_ON_CLOSE` | `false` | Whether closing the palette clears the content. `false` keeps it for the next launch |
| `SAME_ROW_TOLERANCE_PT` | `10` | Difference in top edge treated as the same row when ordering loaded text (pt) |

### Notes

- An alert is shown when no document is open (typing and saving a memo still work).
- The window position is always preserved, regardless of `CLEAR_ON_CLOSE`. Escape also closes the palette.
- Running the script again while the palette is open brings the existing palette to the front instead of opening a second one.
- Clear restarts the script and reopens the palette empty.
- Copy All copies through a temporary text frame, so the clipboard holds an Illustrator object. Other applications may not be able to paste it as plain text.
- Text inside symbols is read from a duplicate placed on a temporary layer; the original symbols are left untouched.

### How it works

The palette runs in a persistent engine (`#targetengine`), and while it is on screen that engine's `app` loses its connection to the document and throws `there is no document`. Every operation that touches the DOM — reading the selected text, reading and writing the clipboard — is therefore delegated to the main engine through BridgeTalk, with the result returned asynchronously.

Illustrator has no `app.system`, so writing to the clipboard goes through a temporary text frame, and reading it means pasting into the document, reading the text, and removing what was pasted. The selection is always cleared before pasting so that characters being edited with the Type tool (a TextRange) are never removed.

### Original idea

Based on the following article by こじらせたクマー, with added features and refactoring.

https://note.com/nice_lotus120/n/n6291a432b30d

### Change log

- v1.1.3 (2026-08-16) : Copying to the clipboard now goes through redraw + menu command for reliability, loading no longer removes characters being edited with the Type tool, internal structure tidied up

### Script info

- Version: v1.1.3
- First release: 2026-06-15
- Last updated: 2026-08-16
- Article: https://note.com/dtp_tranist/n/n41e91e4b1a09
