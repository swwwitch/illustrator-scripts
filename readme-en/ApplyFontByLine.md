# ApplyFontByLine

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyFontByLine.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ApplyFontByLine.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Treats each line (paragraph) of the selected text frames as a font name, looks it up, and applies the matching font to that line.
- Handy for building font specimen sheets: confident matches are applied automatically, and only fuzzy or failed matches go through an interactive picker.
- Frames that still contain unapplied lines get a translucent red marker placed behind them.

### Main Features

- Recursively collects text frames from the selection (including inside groups; locked and hidden items are skipped)
- Staged font matching
  - Forced assignment for specific strings via CUSTOM_MAP (for example Jenson → Adobe Jenson Pro)
  - Exact PostScript name, exact family + style, and exact family name (confident → applied automatically)
  - Partial family match, partial full-name match, and first-word match (3 characters or more) (fuzzy → used as the picker's initial value)
  - When several styles qualify, STYLE_PRIORITY decides (bold → semibold → medium → regular)
  - Matching lowercases the text and strips spaces (half- and full-width) and periods to absorb naming variations
- Phase 1 applies only confident matches, with a progress bar
- Phase 2 resolves the queued lines in a font picker
  - Shows the **Target text** and zooms to the frame (fitting it to 60% of the visible area)
  - The **Search** field filters the family dropdown; pick a **Font** and a **Style**
  - **Apply** / **Skip** / **Quit** (marks all remaining queued lines as unapplied and stops)
  - The same string is never asked twice in one run; the earlier choice is reused
- Creates a red rectangle at 35% opacity behind each frame with unapplied lines, on a layer named "// missing-fonts", then locks that layer (an existing marker layer is removed on every run)
- Lists the unapplied strings in a dialog with a copy-to-clipboard button
- Automatic Japanese / English UI

### Workflow

1. Collect text frames recursively from the selection and build the font index
2. Match each line: apply confident matches, queue everything else
3. Resolve the queue one entry at a time in the font picker (Skip and Quit are recorded as unapplied)
4. Place markers on the frames with unapplied lines and show the list of unapplied strings

### Not Supported

- No open document (alert)
- An empty selection, or a selection containing no text frame (alert)
- Locked or hidden objects (a locked or hidden group is skipped along with its contents)
- Empty lines and whitespace-only lines
- Only the font is changed; text contents and size are left alone
- No live preview while the modal dialog is open (removed as a crash countermeasure)

### Update History

- v1.0.0: Initial release
- v1.1.1: Current version
