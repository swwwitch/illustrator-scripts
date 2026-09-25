# List the document's text, edit it and write it back

[![Direct](https://img.shields.io/badge/Direct%20Link-TextScopeEdit.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextScopeEdit.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextScopeEdit.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Lists the text in the document, including text in symbols, and writes your edits back while keeping the formatting.
You can also review layer, artboard and font names, and export the text and font names to a text file.

<img alt="The Collect and Edit Text dialog" src="../png/ss-1160-1238-144-20260926-013953.png" width="50%" />

### Features

- Text tab: lists the target text; select a row and edit it
  - Text in symbols is listed at the end, marked with ♣, and can be edited the same way
  - Update applies the edit to the document so you can move on to the next row without closing the dialog
- Layer Names and Artboard Names tabs: show the name lists
- Font Names tab: lists the fonts in use, including those in symbols; click a row to select the text that uses it
- Export Text...: writes the text and font names, grouped by artboard, to a text file on the desktop
- Copy Text: copies the full text of every row in the text list to the clipboard

### Usage

1. Run the script.
2. Choose which text to list with Scope and Text to Include.
3. Select a row in the Text List and rewrite it in Edit Text.
4. Click Update to keep editing, or OK to finish.

Type a forced line break with Shift+Enter (shown as `@#` in the edit field).

### Options

| Option | Description |
|---|---|
| Scope | Current Artboard / All Artboards. Include Outside Artboards widens it to the whole document |
| Text to Include | Whether to include layers starting with //, locked or hidden text, and text in symbols |
| Sort | None / By Position (top to bottom, left to right at the same height) / Alphabetical |
| Edit Identical Text Together | Lists identical text as one row and applies the edit to every copy |
| Keep Formatting | Rewrites only the changed characters and keeps character and paragraph formatting. When off, the whole text takes the formatting of its first character |

### Notes

- Editing text in a symbol rewrites the symbol definition. Instances outside the scope change too, and symbol options such as the registration point are reset.
- Edits applied with Update are not reverted by Cancel (use Illustrator's Undo).

### Article (Japanese)

https://note.com/dtp_tranist/n/nb845889dd553

### Update History

- v1.5.0 (2026-09-26)
  - Added Copy Text (merged in from TextExport.jsx)
- v1.4.0 (2026-09-26)
  - Text in symbols can now be edited (added "Text in Symbols" to Text to Include)
  - Text inside locked groups or locked parent layers can now be rewritten
  - Renamed Keep Paragraph Formatting to Keep Formatting; it now rewrites only the changed characters (fixes the first character of line 2 taking line 1's formatting)
  - Added an Update button below the edit field to apply edits without closing the dialog
  - The Font Names tab and exported font names now include fonts used in symbol text
  - Fixed the selection being cleared after export, and font selection matching fonts whose names only partly match
  - Removed the Preview option, and moved the edit options into an Options panel
  - Revised panel and option wording (Canvas tab → Text tab, and others)
- v1.3.6 (2026-04-08)
