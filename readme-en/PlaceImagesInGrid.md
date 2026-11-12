# Place chosen images in a grid inside a rectangle or artboard

[![Direct](https://img.shields.io/badge/Direct%20Link-PlaceImagesInGrid.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/PlaceImagesInGrid.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PlaceImagesInGrid.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Places the images you choose in the file dialog box (one or more) as links and lays them out in a grid inside the selected rectangle. With nothing selected, they are laid out on the active artboard.

After placing, a dialog box lets you adjust the columns, rows, margins and the alignment of the last row while previewing.

### Features

- Choose and place several images at once (as links)
- Target: the selected rectangle, or the active artboard
- Columns and rows are calculated to fill the target as much as possible (Auto), or set by hand (Custom)
- Each image is scaled to fill its cell while keeping its aspect ratio
- Margin between images (Between items) and margin inside the target (Inside target)
- Leftover space is split evenly above and below (or left and right)
- A short last row can be aligned Left or Center
- Shuffle rearranges the images in random order
- Live preview as you type
- Stepper buttons and arrow keys for numeric fields
- Margins are entered in Illustrator's ruler units
- Japanese / English UI

### How to use

1. Select one rectangle to fill with images. To use the artboard, select nothing.
2. Run `PlaceImagesInGrid.jsx`.
3. Choose images in the file dialog box (one or more).
4. In the dialog box that opens after placing, adjust the settings while previewing.
   - Grid: Auto / Custom, Columns, Rows, Last row (Left / Center)
   - Margins: Between items, Inside target
5. Click Shuffle to rearrange the images if you like.
6. Click OK.

### Notes

- If anything other than a single rectangle is selected (several objects, or a non-path object), the script shows a message and stops before the file dialog box opens.
- The rectangle's bounding box (stroke excluded) is used as the frame. The rectangle is kept, and the images are placed just in front of it.
- Images with different aspect ratios do not end up the same size: wide images fit the cell's width and tall images fit its height.
- When you set the rows, the columns follow. Images fill row by row from the top left, so row counts that cannot be made (such as 5 rows for 16 images) are skipped.
- If the margins leave no room, the images are lined up in a row at 100%.
- Cancel removes the placed images.
- Images cannot be placed while the active layer is locked or hidden.

### Update History

- v1.0.0 (2026-11-12) Initial release
