# Generate rules and backgrounds matching the text frames

[![Direct](https://img.shields.io/badge/Direct%20Link-TableMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/TableMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TableMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Generates rules and backgrounds that match the appearance — position, width and line count — of the selected text frame.

The text itself, including tabs, styles and tab stops, is never touched.

### Features

- Modes: rectangle border, no border, or one rectangle per line
- Vertical rules are placed using the existing tab stops

### Usage

1. Select the text frame.
2. Run the script.
3. Choose the mode and the options, then run it.

### Notes

- Because the text is untouched, the rules do not follow later text edits — regenerate them instead.

### Article

https://note.com/dtp_tranist/n/n4eaa14098858

### Update History

- v1.0 (2026-01-24)
- v1.0.2 (2026-09-22): Fixed the stroke width not being converted, so no rules were created, when the ruler unit was H, ft and similar units. Fixed Vertical rules skipping some columns when a paragraph had a tab stop at 0. Internal cleanup
