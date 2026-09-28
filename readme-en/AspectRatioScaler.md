# Resize to Aspect Ratio

[![Direct](https://img.shields.io/badge/Direct%20Link-AspectRatioScaler.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/AspectRatioScaler.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AspectRatioScaler.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Resizes the selected objects to a chosen aspect ratio (Original Ratio, 16:9, 1:1, A4 or Custom). Fix the width or height and set the other side from the ratio, a length or a percentage, or choose Fixed: None (Free) to set width and height separately. With nothing selected, it draws a rectangle at the center of the artboard.

<img alt="Resize to Aspect Ratio dialog" src="../png/ss-914-734-144-20260927-044110.png" width="50%" />

### Key Features

- Aspect ratio: Original Ratio (default), 16:9, 1:1 (Square), A4 (1:1.414), or Custom (enter width:height)
- Orientation: choose Portrait or Landscape with icons
- Reference point: pick the point that stays put with the 9-axis widget (center by default)
- Size: the side chosen under Fixed keeps its original length; the other side follows the ratio, or a length or percentage you enter, which overrides the ratio
- Fixed: None (Free): ignores the ratio and lets you set width and height separately, as a length or percentage (untouched sides stay as they are)
- Resulting ratio: while a ratio other than Custom is chosen, the custom fields show the resulting ratio (switch to Custom to edit from there)
- Live preview on the artboard as you change settings
- Reset: returns the settings to the defaults and restores the size from before the dialog opened
- Numeric fields step to the next whole number (1.5 → 2) with the ∧∨ buttons and arrow keys, to the next multiple of 10 with Shift, and by ±0.1 with Option

### Usage

1. Select objects and run the script (it also runs with nothing selected)
2. Choose the aspect ratio and orientation
3. Under Fixed, choose the side that keeps its original length; optionally set the other side as a length or percentage
4. Optionally set the reference point
5. Click OK to apply, or Cancel to restore

### Options

- **Make Pixel Perfect**: runs Make Pixel Perfect when you click OK (on by default)
- **Add Artboard**: adds an artboard matching the result; the objects stay in place

### Notes

- Sizes use the ruler unit, shown in the panel title (e.g. Size (mm))
- Percentages are relative to each object's original length
- Choosing another ratio, orientation or Fixed side drops the typed length or percentage and returns to the ratio (with Fixed: None (Free), changing the ratio or orientation keeps them)
- With several objects selected, lengths are not shown; the percentage and custom fields show a value only when it is the same for every object
- With nothing selected, the rectangle is 100 mm wide (1000 px when the ruler is in pixels, 200 pt for other units) at 16:9

### Article (Japanese)

https://note.com/dtp_tranist/n/n4a212e6eacf1

### Update History

- v1.0 (20250720): Initial release
- v1.1 (20250721): Added artboard conversion & custom ratio
- v1.2 (20250722): Improved dialog structure, localization, and key input
- v1.5.2 (20260923): Fixed decimals being dropped while typing in numeric fields; added shift/option stepping; revised UI wording
- v1.6.0 (20260927): Two-column dialog layout; "Make Pixel Perfect" now on by default; width and height fields shown together (only the fixed side is editable); added a reference point (9-axis) picker; orientation is now chosen with icons; shows an alert when no document is open; the width is now rounded to the ruler unit when the height is fixed; revised the dialog title, panel names and tooltips
- v1.7.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.8.0 (20260927): The settings from the last OK (ratio, custom ratio, options, orientation, reference point, fixed side) are kept while Illustrator runs; Reset returns to the defaults. Fixed now defaults to Height. Vertically aligned the orientation icons with the reference point
- v1.7.1 (20260927): Added Original Ratio to the ratios (now the default). The fixed side now keeps its original length while the other side takes a length or percentage. Added None (Free) to Fixed (enter width and height separately). The resulting ratio shows in the custom fields, and the size unit appears in the panel title. Added a Reset button that restores the size from before the dialog opened. Revised UI wording
- v1.8.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.8.2 (20260928): The button row is now built with the shared part
- v1.8.2 (20260928): The 3×3 reference point picker now uses the shared part
- v1.8.3 (20260929): Dialog opacity changed to 98%
