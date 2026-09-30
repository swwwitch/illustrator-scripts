# Show and reset the scale of placed images

[![Direct](https://img.shields.io/badge/Direct%20Link-ImageScaler.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/ImageScaler.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageScaler.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Displays and rescales the scale percentage (%) of selected placed images (PlacedItem / RasterItem).
- Calculates actual scale from transformation matrix including rotation/skew, and supports value changes via arrow keys in the input field.

### Main Features

- Calculate actual X/Y scale from transformation matrix
- Dialog with scale % input (prefilled for single selection)
- Apply relative scaling immediately on value change
- Keyboard increments:
  - ↑↓ = to the next whole number (1.5 → 2). Clicking the stepper buttons left of the field works the same way
  - Shift+↑↓ = to the next multiple of 10 (232 → 240)
  - Option(Alt)+↑↓ = ±0.1

### Process Flow

1. Check selection and filter to placed/raster items
2. Prefill scale value if only one item is selected
3. Show dialog, apply changes immediately (app.redraw)
4. OK button closes the dialog

### Update History

- v1.0 (20250816) : Initial version
- v1.1 (20250816) : Added arrow key increment feature
- v1.2 (20250816) : Immediate application of changes (OK closes only), localization support
- v1.3.2 (20260927) : Fixed typing other than the arrow keys re-rounding the value. Added a colon to the field label and removed the extra space in the dialog title
- v1.4.0 (20260927) : Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.4.1 (20260928) : The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.4.2 (20260928) : The button row is now built with the shared part
- v1.4.3 (20260929) : Dialog opacity changed to 98%
- v1.4.4 (20260930) : Fixed an error when running with characters selected by the Type tool
- v1.4.5 (2026-10-01) The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part

### Script info

- Version: v1.4.3
