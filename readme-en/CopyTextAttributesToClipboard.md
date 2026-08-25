# CopyTextAttributesToClipboard

[![Direct](https://img.shields.io/badge/Direct%20Link-CopyTextAttributesToClipboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/CopyTextAttributesToClipboard.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyTextAttributesToClipboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Reads the character and paragraph attributes of the selected text, taking the first character as the reference, and stores them.

They are stored in `$.global.FontClipboard` inside the persistent "FontClipboard" engine, ready for ApplyTextAttributesFromClipboard.jsx to apply.

### Features

- Font (PostScript name, family, style), font size, leading, auto leading
- Tsume, tracking, auto-kerning (value and display name), proportional metrics
- Composition direction and justification (value and display name)

### Usage

1. Select the source text (a partial selection with the Type tool is fine).
2. Run the script.
3. Apply the values with ApplyTextAttributesFromClipboard.jsx.

### Notes

- The first character of the selection is used as the reference.
- The stored values live in `#targetengine "FontClipboard"` and are lost when Illustrator quits.

### Update History

- v1.3.0 (2026-05-21)
