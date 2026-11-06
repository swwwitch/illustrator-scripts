# Arrange objects in a masonry layout

[![Direct](https://img.shields.io/badge/Direct%20Link-MasonryLayout.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/MasonryLayout.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MasonryLayout.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Rearranges the selected objects in a masonry layout, placing each one in the shortest column. The reference is the selected objects or an artboard; for the artboard, the column count and size are chosen to fit inside the margin. You can also match widths to the column and randomize the order. You can check the result in the live preview.

### Main Features

- Packs objects of different heights so the columns stay as even as possible.
- Objects wider than a column span several columns.
- **Reference** sets the area to lay out: **Selected Objects** or **Artboard** (the default).
- With **Selected Objects**, you set the column count and column width. The dialog opens with the first object's width as the column width and as many columns as fit the width of the selection.
- With **Artboard**, fitting takes priority over the column settings. The script picks the column count that keeps the objects as large as possible and scales them down, strokes included, if they do not fit at actual size.
- **Fit to Artboard**, on by default, scales up as well, to fill the artboard. Proportions are kept, so space remains on either the width or the height.
- **Match Width to Column** scales every object to the column width, keeping its proportions. It suits collages of photos in mixed sizes.
- **Rotate 90° to Fit Column** rotates only the objects that would span several columns and fit one column once rotated (off by default). Objects that still would not fit one column stay as they are and span several columns.
- With **Artboard**, **Margin** keeps that much space inside the artboard edges.
- The horizontal and vertical gaps can be set separately. While the link icon is on, the two values stay linked (on by default).
- Choose where to start: from the left or right, and from the top or bottom.
- With **Artboard**, **From Center** is also available. It centers the whole layout on the artboard, horizontally, vertically or both (both by default). Switching back to **Selected Objects** returns to left and top.
- **Order** is **Position** (rows from the top, left to right in each row), **Stacking** (from the front) or **Random** (the default).
- With **Random**, **Shuffle** makes a new order. Changing other values keeps the current order.
- A single undo reverts the result.

### Usage

1. Select two or more objects.
2. Run the script.
3. Choose the **Reference** and adjust the columns, gaps and arrangement.
4. Click **OK** to apply, or **Cancel** to go back.

### Notes

- A group moves as a single object. Locked objects, hidden objects and guides are left out.
- Bounds follow the **Use Preview Bounds** preference.
- With **Match Width to Column** off, **Selected Objects** never changes the size of objects.
- **Match Width to Column** scales strokes by the same ratio. Lines with no width are left as they are.
- **Rotate 90° to Fit Column** is unavailable while **Match Width to Column** is on.
- **Position** order uses the positions before rotation.
- With the default settings, the preview changes the size and position of objects as soon as the dialog opens. **Cancel** restores them. The order is **Random**, so each run gives a different arrangement.
- **Artboard** is the artboard that holds the most selected objects. On a tie, or when none holds any, it is the active artboard.
- With **Artboard**, gaps keep their size. The column count field is disabled and shows the number chosen.
- With **Artboard**, **Column Width** is a reference at actual size, used to decide rotation and the ratio of columns to objects. The whole layout is scaled, so the final column width can differ from the value entered.
- Switching back to **Selected Objects** resets the column count to the number that fits the width of the selection.

### Credits

The layout calculation is ported from [Masonry](https://masonry.desandro.com) v3.3.2 (MIT License) by David DeSandro.

### Update History

- v1.0.0 (2026-11-07): Initial release
- v1.1.0 (2026-11-07): Added Fit to Artboard, Match Width to Column, Margin, and From Center for the artboard
- v1.1.1 (2026-11-07): The dialog now opens with Current Artboard, Fit to Artboard on, From Center both ways, and Random order
- v1.2.0 (2026-11-07): Added Rotate 90° to Fit Column. Renamed the Current Artboard reference to Artboard, which now uses the artboard holding the selected objects. Column Width can now be used with the artboard
