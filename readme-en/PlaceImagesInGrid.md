# Place chosen images in a grid inside a rectangle or artboard

[![Direct](https://img.shields.io/badge/Direct%20Link-PlaceImagesInGrid.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/PlaceImagesInGrid.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PlaceImagesInGrid.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Places the images you choose in the file dialog box (one or more) as links and lays them out in a grid inside the selected rectangle. With nothing selected, they are laid out on the active artboard.

After placing, a dialog box lets you adjust the columns, rows, margins, the alignment of the last row and a mask for each image (corner radius and stroke) while previewing.

### Features

- Choose and place several images at once (as links)
- Target: the selected rectangle, or the active artboard
- Columns and rows are calculated to fill the target as much as possible (Auto), or set by hand (Manual)
- Each image is scaled to fit its cell while keeping its aspect ratio
- Space between images (Gutter) and margin around the area (Inset)
- Leftover space is split evenly above and below (or left and right)
- A short last row can be set to Align Left or Align Center
- Shuffle rearranges the images in random order
- Mask each image (Cell / Image), with corner radius applied as an effect
- Stroke the mask rectangle (weight and color)
- Dialog settings are remembered for the next run
- Save, load and delete settings as presets, with built-in presets (Black 0.3 pt, square / White 1 pt, radius 2 pt / White 1 pt, radius 50 pt)
- Live preview as you type
- Stepper buttons and arrow keys for numeric fields
- Margins are entered in Illustrator's ruler units
- Japanese / English UI

### How to use

1. Select one rectangle to fill with images. To use the artboard, select nothing.
2. Run `PlaceImagesInGrid.jsx`.
3. Choose images in the file dialog box (one or more).
4. In the dialog box that opens after placing, adjust the settings while previewing.
   - Preset: load a saved or built-in preset, save the current settings (save icon), delete the selected preset (trash icon)
   - Grid: Auto / Manual, Columns, Rows, Last row (Align Left / Align Center)
   - Margins: Gutter, Inset
   - Mask shape: None / Cell / Image
   - Mask options: Corner radius, Stroke weight, Stroke (color)
5. Click Shuffle to rearrange the images if you like.
6. Click OK.

### Notes

- If anything other than a single rectangle is selected (several objects, or a non-path object), the script shows a message and stops before the file dialog box opens.
- The rectangle's bounding box (stroke excluded) is used as the frame. The rectangle is kept, and the images are placed just in front of it.
- Images with different aspect ratios do not end up the same size: wide images fit the cell's width and tall images fit its height.
- When you set the rows, the columns follow. Images fill row by row from the top left, so row counts that cannot be made (such as 5 rows for 16 images) are skipped.
- If the margins leave no room, the images are lined up in a row at 100%.
- With Cell, each image is scaled to cover its cell and clipped by the cell rectangle. Parts outside the cell are hidden, and all images end up the same size.
- With Image, each image is clipped at its own size. Nothing is cropped; use it to add rounded corners or a stroke.
- The corner radius is applied to the clip group as Effect > Stylize > Round Corners, so you can change it later in the Appearance panel.
- The corner radius and the stroke are applied only when their checkboxes are on. They are unavailable when Mask shape is None.
- The stroke is applied to the clipping path. The weight is entered in the stroke unit set in Preferences. Click the Stroke swatch to choose a color in the standard color picker.
- Settings are saved when you click OK and used as the defaults next time (the column count carries over only in Manual mode).
- A preset stores the columns, rows, margins, last row, mask shape, corner radius and stroke. Presets are saved to `Folder.userData/illustrator-scripts/PlaceImagesInGridPresets.json` and survive a restart.
- Built-in presets change only the mask stroke and corner radius (other settings stay as they are), and switch Mask shape from None to Cell. They cannot be deleted, and their names cannot be used for your own presets.
- Cancel removes the placed images.
- Images cannot be placed while the active layer is locked or hidden.

### Article

https://note.com/dtp_tranist/n/n3b15f8368c27

### Update History

- v1.2.0 (2026-11-12) Added presets (save, load, delete) and three built-in presets (Black 0.3 pt, square / White 1 pt, radius 2 pt / White 1 pt, radius 50 pt)
- v1.1.2 (2026-11-12) Revised the dialog box wording (title "Place Images in Grid", Gutter, Inset, Cell / Image, Manual, Align Left / Align Center). Removed the Embed images option. Adjusted the dialog box layout (narrower labels, stacked last-row options, a checkbox on Stroke weight placed above Stroke)
- v1.1.1 (2026-11-12) Rearranged the dialog box into two columns (grid and margins on the left, mask and options on the right) and aligned the label widths. Stacked the mask modes vertically and added a checkbox for the corner radius
- v1.1.0 (2026-11-12) Added masks (cell shape / item shape), corner radius (effect), a stroke on the mask (color and weight), an option to embed the images, and remembered settings
- v1.0.0 (2026-11-12) Initial release
