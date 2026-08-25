# InsertNewAreaText


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--areatext.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-areatext.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewAreaText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Creates a rectangle of a given size at the center of the current view, converts it into an area text,
and fills it with sample text. Font, size, leading and justification are applied, and the created area
text is left selected. A demo script for building samples and mockups.

### Features

- Creates an `AREA_WIDTH` × `AREA_HEIGHT` rectangle (220 × 110 by default) at the center of the view
- Converts it into an area text and pours in `SAMPLE_TEXT`
- Tries `FONT_CANDIDATES` (Shuei Gothic L → Hiragino Sans W3 → Source Han Sans) in order
- Applies 12pt font size, 18pt leading and 100% horizontal / vertical scaling
- Sets justification to full justify with the last line flush left
- Applies proportional metrics (`APPLY_PROPORTIONAL_METRICS`)
- Selects the created area text only

### How it works

1. Check that a document is open
2. Find the first unlocked, visible layer
3. Create a rectangle at the center of the view
4. Convert it into an area text and pour in the sample text
5. Apply the text style and the font
6. Select the created area text

### Notes

- The area size, sample text, font candidates and text style live in the User Settings block at the top of the script.
- `TEXT_STYLE.justification` takes a `Justification` key name as a string (`FULLJUSTIFYLASTLINELEFT`, `CENTER`, …). An unknown key raises an error.
- If no editable layer exists, the script exits quietly without an alert.
- The sample text uses `\r` (carriage return) for line breaks: from ExtendScript, `\r` — not `\n` — separates paragraphs.

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### Update history

- v1.0 (20250813) : Initial release
