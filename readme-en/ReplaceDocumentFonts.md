# Replace the fonts in use with another font

[![Direct](https://img.shields.io/badge/Direct%20Link-ReplaceDocumentFonts.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ReplaceDocumentFonts.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceDocumentFonts.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that lists the fonts used in the document (or the selection) by family and style, and replaces the selected ones with another font in a single pass.
- The dialog has two columns: "Source Fonts" (multiple selection) and "Target Font". Below each list is a panel of display or replace options.
- Each style row shows, in parentheses, how many text objects use that font.
- The dialog stays open after a replacement and the list is rebuilt automatically, so you can keep replacing.

<img alt="The Replace Document Fonts dialog" src="../png/ss-1140-1190-144-20260929-205620.png" width="70%" />

### Main Features

#### Scope

- Switch between Entire Document and Selection Only
- Selection Only targets the text selected when the script starts (including inside groups; while editing text, that story)
- Running the script with text selected opens in Selection Only, otherwise in Entire Document. Selection Only is unavailable when no text is selected

#### Lists

- **Source Fonts** (multiple selection): selecting a family row selects every style in that family
- **Target Font** (single selection): family rows cannot be selected (selecting one clears the selection)
- Counts are per text object: a font used several times inside the same text object is counted once
- Family rows show the total count of their styles
- A family with only one style is shown on a single row ("family + style") with no header

#### Display (left panel)

- **Show font names as PostScript names**: drop the family / style nesting and list one row per font by PostScript name
- **Fit list width to font names**: when on, the lists widen to the longest font name so nothing is cut off; when off, they use a compact fixed width
- **Composite fonts only**: list only composite fonts in both the source and target lists; Replace All then affects composite fonts only. Unavailable when no composite font is in use (always starts off)
- **Sort**: choose Name, Most Used, or Least Used. Families are ordered by the total count of their styles; ties fall back to the name

#### Replace Options (right panel)

- **Show unused font styles**: also list, without a count, the styles of each family that the document does not use (always starts off)
- **Also replace in styles**: switch character and paragraph styles that use a source font to the target font (Entire Document only)

#### Buttons

- **Replace Fonts**: replace the selected source fonts with the target font
- **Replace All**: unify every font in use on the target font. With no target selected, everything is unified on the first source font instead
- After a replacement the list is rebuilt and the selection is restored by font name (fonts no longer in use drop off the list)

#### Keyboard Shortcuts

| Key | Action |
|---|---|
| Option+D | Switch to Entire Document |
| Option+S | Switch to Selection Only |
| Option+P | Toggle "Show font names as PostScript names" |
| Option+L | Toggle "Fit list width to font names" |
| Option+C | Toggle "Composite fonts only" |
| Option+Tab | Move focus between the source and target lists |

#### Other

- Sort, "Show font names as PostScript names", "Fit list width to font names", and "Also replace in styles" carry over to the next run
- Tooltips on every option; automatic Japanese / English UI
- The style-row indent, the initial state of each option, and the list height / width can be changed in the "User settings" and "Layout" blocks at the top of the script

### How to Use

1. Open a document and run the script. To replace fonts in part of the text only, select that text first.
2. Pick the fonts to replace in "Source Fonts" on the left (multiple selection; selecting a family row takes the whole family).
3. Pick the replacement in "Target Font" on the right.
4. Click "Replace Fonts". To unify every font in use on one font, click "Replace All" instead.
5. Keep replacing as needed, then click "Close".

### Workflow

1. Walk the text objects in the scope (entire document or selection) and collect the fonts in use, grouped by family
2. Build the list (family headers + style rows, or one row per font in PostScript-name mode) and fill both list boxes
3. On replacement, swap the font on the matching ranges, skipping locked and hidden text (and in character / paragraph styles when that option is on)
4. Collect the fonts again, refresh the lists, and redraw

### Not Supported

- No open document, or no fonts in use found
- Locked / hidden text (it is listed and counted, but never replaced)
- Outlined text and text inside placed images
- A target font that is not installed (an alert is shown and the run is aborted)

### Article

https://note.com/dtp_tranist/n/ncc9330ba1f7d (Japanese)

### Update History

- v2.2.0 (2026-09-29) Added "Composite fonts only" to the Display panel (limits the source and target lists to composite fonts; Option+C)
- v2.1.1 (2026-09-29) Removed highlighting of the text matching the source selection. Running the script with text selected now opens Scope in Selection Only (Scope is no longer remembered). Added "Fit list width to font names" (off: compact fixed width). "Show unused font styles" (formerly "Show unused styles") is no longer remembered and always starts off. "Show PostScript names" renamed to "Show font names as PostScript names". Sort is now a popup menu instead of radio buttons. The left panel is renamed "Display". Alert messages reworded and tooltips expanded. Added keyboard shortcuts (Option+D / S / P / L / Tab). Close moved to the far left
- v2.1.0 (2026-09-29) Added Scope (entire document / selection only), Sort (name / most used / least used), "Show unused styles" for the target list, and "Also replace in styles". "Show PostScript names" and Sort now sit in an Options panel on the left, and the two target options in a Replace Options panel on the right. Family rows show the total count of their styles. Settings are remembered. The list no longer jumps when a family row is selected. Dialog title changed to "Replace Document Fonts"
- v2.0.3 (2026-09-29) Dialog opacity changed to 98%
- v2.0.2 (2026-09-28) The button row is now built with the shared part; Cancel moved to the right
- v2.0.1 (2026-09-28) The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v2.0.0 (2026-09-17) Added family / style nesting with a per-style usage count, "Show PostScript names", highlighting of the text matching the source selection, and "Replace All". Widened the dialog. Applied the house rules (user-settings / layout blocks, nested LABELS, JSDoc, tooltips)
- v1.0.0 (2025-03-29) Public release
