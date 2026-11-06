# Arrange objects in a masonry layout

[![Direct](https://img.shields.io/badge/Direct%20Link-MasonryLayout.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/MasonryLayout.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MasonryLayout.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Rearranges the selected objects in a masonry layout, placing each one in the shortest column. The reference is the selected objects or the current artboard; for the artboard, the column count is chosen to fit, scaling down if needed. You can check the result in the live preview.

### Main Features

- Packs objects of different heights so the columns stay as even as possible.
- Objects wider than a column span several columns.
- **Reference** sets the area to lay out: **Selected Objects** or **Current Artboard**.
- With **Selected Objects**, you set the column count and column width. The dialog opens with the first object's width as the column width and as many columns as fit the width of the selection.
- With **Current Artboard**, fitting takes priority over the column settings. The script picks the column count that keeps the objects as large as possible and scales them down, strokes included, if they do not fit at actual size. It never scales up.
- The horizontal and vertical gaps can be set separately. While the link icon is on, the two values stay linked (on by default).
- Choose where to start: from the left or right, and from the top or bottom.
- **Order** is **Position** (rows from the top, left to right in each row), **Stacking** (from the front) or **Random**.
- With **Random**, **Shuffle** makes a new order. Changing other values keeps the current order.
- A single undo reverts the result.

### Usage

1. Select two or more objects.
2. Run the script.
3. Choose the **Reference** and adjust the columns, gaps and arrangement.
4. Click **OK** to apply, or **Cancel** to go back.

### Notes

- A group moves as a single object. Locked and hidden objects are left out.
- Bounds follow the **Use Preview Bounds** preference.
- **Selected Objects** never changes the size of objects.
- With **Current Artboard**, gaps are not scaled. The column count and column width fields are disabled, and the column count field shows the number chosen. No margin is kept from the artboard edges.
- Switching back to **Selected Objects** resets the column count to the number that fits the width of the selection.

### Credits

The layout calculation is ported from [Masonry](https://masonry.desandro.com) v3.3.2 (MIT License) by David DeSandro.

### Update History

- v1.0.0 (2026-11-07): Initial release
