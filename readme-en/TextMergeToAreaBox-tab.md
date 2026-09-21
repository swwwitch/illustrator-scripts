# Merge several text items into one area text box

[![Direct](https://img.shields.io/badge/Direct%20Link-TextMergeToAreaBox--tab.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextMergeToAreaBox-tab.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextMergeToAreaBox-tab.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Notes

https://note.com/dtp_tranist/n/ne8d31278c266

### Overview

- Gathers scattered text items line by line and merges them into a single area text box.
- Items in the same line are joined with tabs, and each line ends with a line break. Useful for rebuilding tables.
- Inherits font, size, and leading from the original text.

### Key Features

- Groups items with close Y positions into one line and joins them left to right with tabs
- Breaks after every line (keeps the rows as they are)
- Replaces line breaks inside an item with "〓" so they do not mix with the row breaks
- Applies "Soft_v2" kinsoku (falls back to "Soft" on versions without it) and justify with last line left to the whole text
- Extends the frame downward when the last line overflows
- With a single line, outputs left-aligned text instead of area text
- Replaces and deletes original text items

### Workflow

1. Sort selected text items from top to bottom and group them into lines
2. Join each line left to right with tabs
3. Create area text from the overall bounding box (one character narrower) and inherit font, size, and leading
4. Delete original items

### Update History

- v1.0 (2025-07-18): Initial release
- v1.1 (2025-07-19): Added support for single line text, set kinsoku rules
- v1.2 (2025-07-20): Added handling for line breaks after English words
- v1.2.2 (2026-09-21): Lines are now simply joined with line breaks (no more blank paragraphs after sentence ends or leading spaces); changed kinsoku to "Soft_v2" (falls back to "Soft" where unavailable); the frame is extended downward when the last line overflows; narrow selections no longer fail; the bounding box no longer depends on the selection; single-line output is also set to horizontal; added a guard when no document is open; code cleanup

### Script info

- Version: v1.2.2
