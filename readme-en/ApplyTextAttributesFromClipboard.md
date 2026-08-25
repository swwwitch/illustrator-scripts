# ApplyTextAttributesFromClipboard

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyTextAttributesFromClipboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ApplyTextAttributesFromClipboard.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyTextAttributesFromClipboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Applies the text attributes saved by CopyTextAttributesToClipboard.jsx to the current text selection.

The values are read from `$.global.FontClipboard` in the persistent "FontClipboard" engine, so the two scripts together work as copy-and-paste for formatting.

### Features

- Four panels: font and size / leading, kerning-related settings, paragraph attributes, and fill plus graphic style
- A checkbox per attribute decides whether it is applied (only font is on by default)
- The fill-and-graphic-style panel is exclusive (none / fill / graphic style)

### Usage

1. Copy the attributes with CopyTextAttributesToClipboard.jsx.
2. Select the target text (a partial selection with the Type tool is fine).
3. Run the script, tick the attributes to apply, and click OK.

### Notes

- With a partial selection in text-edit mode, only that range is changed.
- With several text frames selected, the same attributes are applied to each of them.
- The values travel through `#targetengine "FontClipboard"`, so they are lost when Illustrator quits.

### Update History

- v1.3.1 (2026-05-21)
