# CopyAsPngLikeFigmaWithDialog

[![Direct](https://img.shields.io/badge/Direct%20Link-CopyAsPngLikeFigmaWithDialog.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/CopyAsPngLikeFigmaWithDialog.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyAsPngLikeFigmaWithDialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that rasterizes selected objects at high resolution and copies them as a PNG-like bitmap to the clipboard.
- Allows configuring resolution, background color, anti-aliasing, and margin via a dialog at runtime.

### Main Features

- Selectable resolution (dpi) from 72 to 1200
- Choose background color (transparent, white, black)
- Toggle anti-aliasing on/off
- Set margin
- Japanese and English UI support

### Process Flow

1. Select target objects
2. Configure settings via dialog
3. Duplicate to temporary layer, rasterize and resize
4. Copy to clipboard
5. Delete temporary objects and restore selection

### Update History

- v1.0.0 (20250502): Initial version
- v1.0.1 (20250603): Refined labels and scaling adjustments
