# InsertNewTextELong


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--text--e--long.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-text-e-long.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewTextELong.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Creates a single, longer point text with a Latin font at the center of the current view, centers the
paragraph, and leaves it selected. A demo script for building samples and mockups.

### Features

- Creates a point text at the center of the view
- Applies the `TEXT_CONTENTS` string (`Design with clarity, build with intent.` by default) and `FONT_SIZE` (12pt)
- Tries `FONT_CANDIDATES` in order and applies the first font found
- Sets the paragraph to centered
- Selects the created text only

### How it works

1. Check that a document is open
2. Clear the selection
3. Get the center of the view
4. Create a text frame and set its contents
5. Set the font first, then the font size and justification
6. Move it to the center and select it

### Notes

- The difference from [InsertNewTextE](InsertNewTextE.md) is the length of the string and the positioning: this one moves the frame using `position` with `width` / `height`.
- The string, font size and font candidates live in the User Settings block at the top of the script.
- If none of the font candidates is available, an alert is shown and the text is created with the current font.
- For Japanese, use [InsertNewTextJ](InsertNewTextJ.md).

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### Update history

- v1.1 (20250813) : Externalized settings and font candidates, split into helpers, added centered justification
- v1.0 (20250401) : Initial release
