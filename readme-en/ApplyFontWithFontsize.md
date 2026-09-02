# ApplyFontWithFontsize

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyFontWithFontsize.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ApplyFontWithFontsize.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyFontWithFontsize.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Treats each line (paragraph) of the selected text frames as a "font name (plus size and leading)" spec, looks the font up, and applies it to that line.
- Handy for building font specimen sheets: confident matches are applied automatically, and only fuzzy or failed matches go through an interactive picker.
- Frames that still contain unapplied lines get a translucent red marker placed behind them.

### Line Format

A line can carry the font name alone, or the name followed by a size and a leading.

| Example line | What gets applied |
| --- | --- |
| `Hiragino Kaku Gothic W3` | Font only |
| `Hiragino Kaku Gothic W3 12pt` | Font and size |
| `Hiragino Kaku Gothic W3 12pt↓16pt` | Font, size and leading |
| `Helvetica Neue, 14` | Font and size |

- The name and the size are separated by a space (half- or full-width), a tab, or one of "、" "," "/"
- The size and the leading accept those separators plus "↓"
- Sizes and leading take pt / px / Q (1Q = 0.25mm); a bare number is read as pt
- A size is always anchored by an explicit unit or by a "、" "," "/" separator, so trailing digits in a font name (such as `DIN 2014`) are never mistaken for one
- The leading is not set as an absolute value: leading ÷ size becomes the auto-leading percentage (12pt↓16pt gives about 133.3%)
- Because of that, a line that sets only a size still changes the leading of a paragraph using auto leading, since it is recalculated as size × the auto-leading percentage

### Main Features

- Recursively collects text frames from the selection (including inside groups; locked and hidden items are skipped)
- Splits off the size and leading written at the end of a line, and matches on the name alone
- Staged font matching
  - Forced assignment for specific strings via CUSTOM_FONT_MAP (for example Jenson → Adobe Jenson Pro)
  - Exact PostScript name, exact family + style, and exact family name (confident → applied automatically)
  - Partial family match, partial full-name match, and first-word match (3 characters or more) (fuzzy → used as the picker's initial value)
  - When several styles qualify, STYLE_PRIORITY decides (bold → semibold → medium → regular)
  - Matching lowercases the text and strips spaces (half- and full-width) and periods to absorb naming variations
- Phase 1 applies only confident matches, with a progress bar
- Phase 2 resolves the queued lines in a font picker
  - Shows the **Target text** and zooms to the frame (fitting it to 60% of the visible area)
  - The **Search** field filters the family dropdown; pick a **Font** and a **Style**
  - **Apply** / **Skip** / **Quit** (marks all remaining queued lines as unapplied and stops)
  - The same font name is never asked twice in one run; the earlier choice is reused (the size and leading still come from each line)
  - A failed apply is not remembered, so the next line with that name is asked again
  - Once the queue is done, the zoom and view position moved by the picker are restored
- Creates a red rectangle at 35% opacity behind each frame with unapplied lines, on a layer named "// missing-fonts", then locks that layer (an existing marker layer is removed on every run)
- Lists the unapplied strings in a dialog with a copy-to-clipboard button
- Automatic Japanese / English UI

### Workflow

1. Collect text frames recursively from the selection and build the font index
2. Split off the size and leading, match on the name: apply confident matches, queue everything else
3. Resolve the queue one entry at a time in the font picker (Skip and Quit are recorded as unapplied), then restore the view
4. Place markers on the frames with unapplied lines and show the list of unapplied strings

### Not Supported

- No open document (alert)
- An empty selection, or a selection containing no text frame (alert)
- Locked or hidden objects (a locked or hidden group is skipped along with its contents)
- Empty lines and whitespace-only lines
- Text contents are never changed; the size and leading change only when the line carries them
- No live preview while the modal dialog is open (removed as a crash countermeasure)

### Article

- [Applying fonts by font name (Japanese)](https://note.com/dtp_tranist/n/n33d152e73f35)

### Update History

- v1.0.0: Initial release
- v1.1.1
- v1.3.5: Merged ApplyFontByLine.jsx, added support for sizes and leading, and renamed to ApplyFontWithFontsize.jsx (current version)
