# InsertNewTextE


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--text--e.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-text-e.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewTextE.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Creates a single point text with a Latin font at the center of the current view and leaves it selected.
A demo script for building samples and mockups.

### Features

- Creates a point text at the center of the view
- Applies the `TEXT_CONTENTS` string and `FONT_SIZE` (12pt by default)
- Tries `FONT_CANDIDATES` in order and applies the first font found
- Sets the paragraph to centered
- Aligns the center of the visible bounds with the center of the view
- Selects the created text only

### How it works

1. Check that a document is open
2. Clear the selection
3. Create a text frame and set its contents
4. Apply the font candidates, font size and justification
5. Align the visible center with the center of the view
6. Select the created text

### Notes

- The string, font size and font candidates live in the User Settings block at the top of the script.
- `visibleBounds` is used instead of `position` because `position` is baseline-based and would leave the text visually off-center.
- If none of the font candidates is available, the text is created with the current document font and no alert is shown.
- For Japanese, use [InsertNewTextJ](InsertNewTextJ.md); for a longer Latin string, use [InsertNewTextELong](InsertNewTextELong.md).

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### Update history

- v1.1 (20260702) : Switched to centering on the visible bounds
- v1.0 (20250401) : Initial release
