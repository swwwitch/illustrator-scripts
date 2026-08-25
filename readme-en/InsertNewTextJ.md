# InsertNewTextJ


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--text--j.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-text-j.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewTextJ.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Creates a single point text with a Japanese font at the center of the current view, centers the paragraph,
and leaves it selected. A demo script for building samples and mockups.

### Features

- Creates a point text at the center of the view
- Applies the `TEXT_CONTENTS` string and `FONT_SIZE` (12pt by default)
- Tries `FONT_CANDIDATES` (Shuei Maru Gothic → Hiragino Sans → Source Han Sans) in order
- Resets horizontal and vertical character scaling to 100%
- Sets the paragraph to centered
- Selects the created text only

### How it works

1. Check that a document is open
2. Clear the selection
3. Get the center of the view
4. Create a text frame and set its contents
5. Set the font first, then the font size, scaling and justification
6. Move it to the center and select it

### Notes

- The string, font size and font candidates live in the User Settings block at the top of the script.
- The font is applied before the size because changing the font afterwards can recompose the text and undo the size.
- If none of the font candidates is available, an alert is shown and the text is created with the current font.
- Positioning uses `position` (baseline-based) together with `width` / `height`, so some glyphs end up slightly off the visual center. See [InsertNewTextE](InsertNewTextE.md) for the visible-bounds approach.
- For Latin text, use [InsertNewTextE](InsertNewTextE.md) or [InsertNewTextELong](InsertNewTextELong.md).

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### Update history

- v1.1 (20250813) : Externalized settings and font candidates, split into helpers, added centered justification
- v1.0 (20250401) : Initial release
