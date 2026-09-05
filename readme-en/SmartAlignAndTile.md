# Tile stacked objects horizontally or vertically

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAlignAndTile.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartAlignAndTile.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Redistributes stacked objects along the horizontal or vertical axis at the spacing you specify. Set a row or column count to tile them.
- Detects the key object set in Illustrator and keeps it in place while the rest are rearranged.
- Every change in the dialog updates the preview, and the preview never pollutes the Undo history.
- Updated: 2026-09-06
- Merges `SmartAlignAndTile-yoko` (horizontal) and `SmartAlignAndTile-tate` (vertical) into one script.

### Main Features

- Choose the tiling direction: horizontal or vertical
- Split into rows (horizontal) or columns (vertical)
- Separate horizontal and vertical spacing ("Link" applies the horizontal value to both)
- Follows the ruler unit (mm, pt, px, Q/H, …)
- Grid layout: cells sized to the largest object, so every cell shares the same width and height
- Vertical alignment (top / middle / bottom / none) and horizontal alignment (left / center / right / none)
- Anchor to the key object (detected automatically; the checkbox is dimmed when none is found)
- Random arrangement, keeping the top-left corner of the whole block in place
- Up/Down keys step the numeric fields (Shift+Up/Down snaps to multiples of 10); the row/column count is clamped to 1 or more
- Undo-safe preview and single-step Undo after confirming

### How to Use

1. Select the objects to rearrange. To pin one of them, click it again while the selection is active so Illustrator makes it the key object.
2. Run the script.
3. Set the direction, row (column) count, spacing, alignment and options. The preview updates as you go.
4. Click OK. A single Undo (⌘Z) reverts the whole result.

### Options

- **Direction**: Horizontal lays objects out left to right and wraps by row count; Vertical lays them top to bottom and wraps by column count.
- **Align (vertical)**: snaps to top, middle or bottom. "None" leaves the vertical position alone and only applies the row/column offset.
- **Align (horizontal)**: snaps to left, center or right. "None" leaves the horizontal position alone.
- **Anchor to key object**: makes the key object the origin of the layout. It stays where it is, and the other objects follow to its right (horizontal) or below it (vertical). Dimmed when no key object can be detected.
- **Use preview bounds**: ON uses visibleBounds (stroke and effects included), OFF uses geometricBounds. Illustrator's preference is toggled while the dialog is open and restored on both OK and Cancel.
- **Grid**: places objects on a grid whose cell matches the largest object. Turning it on defaults both axes to centered.
- **Random**: shuffles the order. The top-left corner of the block stays put.
- **Link**: applies the horizontal spacing to the vertical one (the V field is dimmed while it is on).

### Notes

- The alignment along the tiling direction (horizontal when tiling horizontally, vertical when tiling vertically) only matters once the cell is larger than the object, so it is dimmed unless Grid is on.
- The key object is detected by running the four align commands and finding the object that never moves. Objects shift briefly during the probe, and the attempt is left in the Undo history.
- No key object is reported when only one object is selected, or when several candidates remain.
- When the item count is not divisible by the row/column count, the remainder is spread over the earlier lanes, so the requested number of rows/columns is always used.

### Original Idea

John Wundes
Distribute Stacked Objects v1.1
https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design
https://gorolib.blog.jp/archives/77282974.html

### note

- [【Illustrator】複数のオブジェクトを整列・タイル配置するスクリプト updated｜DTP Transit 別館](https://note.com/dtp_tranist/n/nf426908d8bcd)

### Update History

- v2.0.0 (2026-09-06): Merged `SmartAlignAndTile-yoko` (v1.7.1) and `SmartAlignAndTile-tate` (v1.8) into one script with a direction switch. The key object now acts as the origin of the layout, "None" was added to both alignment rows, and the lane band is unified on the largest item size. Also fixed the Shift+Down snap, the row/column count clamping, the lane distribution (the requested count is always used) and the preference restore.

---

### Script info

- Version: v2.0.0
