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

### Workflow
After running this script, the expected workflow is:

1. Rearrange artboards using Illustrator's standard features
2. Adjust mask path size using the ResizeClipMask script
https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ResizeClipMask.jsx

### Change Log

- v1.0 (20250710): Initial version
- v1.1 (20250710): Added artboard conversion and options
- v1.2 (20250710): Added shape variations and custom settings
- v1.3 (20250710): Minor adjustments
- v1.4 (20250710): Adjusted aspect ratio options