# InsertNewRectangle5Times


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--rectangle--5times.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-rectangle-5times.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewRectangle5Times.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Creates five black squares with random opacity near the center of the current view, places them so that
they do not overlap, and applies Convert to Shape and Make Pixel Perfect. A demo script for building
samples and mockups.

### Features

- Creates `RECT_COUNT` squares (5 by default) of `RECT_SIZE` (100 by default)
- Fills them with K100 in a CMYK document, or R0/G0/B0 in an RGB document
- Sets a random opacity between `OPACITY_MIN` and `OPACITY_MAX` (30–100% by default)
- Removes an existing rectangle of the same size sitting at the center of the view
- Repositions the squares at random until they are at least `padding` (6 by default) apart
- Widens the search range step by step when they do not fit (up to `maxScaleFactor`)
- Selects the five squares, then runs Convert to Shape and Make Pixel Perfect

### How it works

1. Check that a document is open
2. Resolve the target layer (switching the active layer if needed)
3. Remove an existing rectangle left at the center of the view
4. Create five squares scattered near the center and set fill, stroke and opacity
5. Reposition them at random until nothing overlaps
6. Select the five squares and run Convert to Shape and Make Pixel Perfect

### Notes

- The count, size, opacity range and overlap-avoidance settings live in the User Settings block at the top of the script.
- Step 3 deletes rectangles of the same size at the same position as the center square. It is meant for re-running right after [InsertNewRectangle](InsertNewRectangle.md), but any existing object matching those conditions is deleted too.
- If a square cannot be placed within `attemptsPerItem` tries, the search range widens by one step. If it still does not fit at `maxScaleFactor`, the squares stay where the last attempt left them.
- Because random values are used, the layout and opacities differ on every run.
- To create a single square, use [InsertNewRectangle](InsertNewRectangle.md).

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### Update history

- v1.3 (20260305) : Place the squares with overlap avoidance
- v1.2 (20250713) : Function cleanup and header updates
- v1.1 (20250511) : Comment cleanup and logic improvements
- v1.0 (20250401) : Initial release
