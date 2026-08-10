# ImageScaler

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
  - ↑↓ = ±1
  - Shift+↑↓ = ±10 (snap to multiples of 10)
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

### Script info

- Version: v1.3
