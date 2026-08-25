# InsertNewRectangle


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--rectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-rectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Creates a single black square at the center of the current view, selects it, and applies Convert to Shape
and Make Pixel Perfect. A demo script for building samples and mockups.

### Features

- Creates a `RECT_SIZE` square (100 by default) at the center of the view
- Fills it with K100 in a CMYK document, or R0/G0/B0 in an RGB document
- Removes the stroke
- Switches to the first unlocked, visible layer when the active layer is locked or hidden
- Selects the created square only, then runs Convert to Shape and Make Pixel Perfect

### How it works

1. Check that a document is open
2. Resolve the target layer (switching the active layer if needed)
3. Get the center of the view
4. Create the square and set its fill and stroke
5. Select the created square only
6. Run Convert to Shape and Make Pixel Perfect

### Notes

- The square size is `RECT_SIZE` in the User Settings block at the top of the script.
- If no unlocked, visible layer exists, an alert is shown and the script stops.
- Make Pixel Perfect rounds coordinates to whole pixels, so the square can end up to about 0.5pt off the exact center depending on the zoom level.
- To create five at once, use [InsertNewRectangle5Times](InsertNewRectangle5Times.md).

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### Update history

- v1.2 (20250713) : Split into functions, header cleanup
- v1.1 (20250511) : Comment cleanup and logic improvements
- v1.0 (20250401) : Initial release
