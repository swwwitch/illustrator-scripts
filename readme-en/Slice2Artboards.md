# Slice2Artboards.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-Slice2Artboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/Slice2Artboards.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Splits selected image or objects into grid pieces using specified rows and columns, masking each piece with a rectangle.
- Each piece can be converted into an artboard automatically.
- Useful for imposition, puzzle layouts, and creating multiple artboards.

![](https://www.dtp-transit.jp/images/ss-826-916-72-20250711-004809.png)

### Features

- Grid splitting with row/column settings
- Offset-based size adjustments
- Aspect ratio presets (A4, Square, US Letter, 16:9, 8:9, Custom)
- Automatic conversion to artboards, naming, zero padding
- Margin setting

### Workflow

1. Configure grid, aspect ratio, and options in the dialog
2. Generate grid masks on execution
3. Optionally add and rename artboards
4. Delete original artwork (optional)

### After running

After running this script, the expected workflow is:

1. Rearrange artboards using Illustrator's standard features
2. Adjust mask path size using the ResizeClipMask script
https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ResizeClipMask.jsx

### Notes

- The offset and margin fields show the ruler unit, but **the values are treated as points**. With the ruler set to mm, "10 mm" still applies 10 pt.
- Entering a non-numeric value for the rows or columns stops the run with a message. Setting just one of them to 0 derives it from the aspect ratio of the selection.
- With "Convert to Artboards" enabled, the artboards that existed before the split are removed.
- A multiple selection is grouped and converted to a symbol before being processed.
- Illustrator has no undo-grouping API, so undoing this run takes multiple steps.

### Change Log

- v1.0 (20250710): Initial version
- v1.1 (20250710): Added artboard conversion and options
- v1.2 (20250710): Added shape variations and custom settings
- v1.3 (20250710): Minor adjustments
- v1.4 (20250710): Adjusted aspect ratio options
- v1.5 (20260727): Removed the ineffective `app.undoGroup` assignment (Illustrator has no such property, so the run was never grouped into one undo step)
- v1.6 (20260727): Removed the unused legacy offset helpers (`createOffsetEffectXML` / `applyOffsetPathToSelection`), which were never called and would have thrown on an out-of-scope reference (offsetting is still done by adjusting the rectangle bounds)