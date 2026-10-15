# Color full-width and half-width alphanumerics and fixed-width glyphs differently

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorNumberByWidth.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ColorNumberByWidth.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorNumberByWidth.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Colors full-width alphanumerics and "！？" red, half-width alphanumerics and "!?" blue, and characters set to fixed-width glyphs green in the selected text.
- Useful for spotting mixed full-width and half-width characters and fixed-width glyph settings at a glance.

### Main Features

- Target characters
  - Full-width: ０–９, Ａ–Ｚ, ａ–ｚ, ！, ？
  - Half-width: 0–9, A–Z, a–z, !, ?
  - Fixed-width glyphs: characters set to the OpenType full-, half-, third- or quarter-width glyphs (not only alphanumerics)
- Fixed-width glyphs take priority over full-width and half-width.
- Click a swatch in the dialog to choose the color with Illustrator's standard color picker.
- Colors are created in the document's color mode (RGB / CMYK).

### Usage

1. Select the text (a range selected with the Type tool also works)
2. Run the script
3. Click the Full-width, Half-width or Fixed-width Glyphs swatch to change the color if needed
4. Click OK

### Notes

- With the Selection tool, the whole text is processed, including text inside groups.
- Other symbols (＃, ％, #, etc.) are not colored.
- Change the initial colors with `FULL_WIDTH_RGB` and the other values at the top of the script. Chosen colors are not remembered for the next run.

### Script info

- Version: v1.1.0

### Update History

- v1.1.0 (2026-10-16) Added coloring characters set to fixed-width glyphs (full-, half-, third- and quarter-width) green. Changed the default half-width color to #0099ff (C100 M40 in CMYK)
- v1.0.0 (2026-10-16) Initial release
