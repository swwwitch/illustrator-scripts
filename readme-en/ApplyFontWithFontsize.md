# ApplyFontWithFontsize

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyFontWithFontsize.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ApplyFontWithFontsize.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyFontWithFontsize.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Reads each line of the selected text frames as a "font name plus size (plus leading)" spec and applies it line by line.

The format puts the size after the font name and the leading after that, as in "Hiragino Kaku Gothic W3 12pt↓16pt".

### Features

- Looks up the font using only the font-name part, then applies the size and leading written alongside it
- Text frames inside groups are handled recursively (locked and hidden ones are skipped)
- Matches on PostScript name, family plus style, or family alone

### Usage

1. Prepare a text frame with "font name size↓leading" on each line.
2. Select that text frame.
3. Run the script.

### Notes

- Without a size and leading, it behaves the same as ApplyFontByLine.jsx.
- Lines whose font cannot be found are left unchanged.

### Update History

- v1.3.3
