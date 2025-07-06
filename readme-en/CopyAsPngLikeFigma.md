# CopyAsPngLikeFigma.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-CopyAsPngLikeFigma.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/CopyAsPngLikeFigma.jsx)

[![Direct](https://img.shields.io/badge/Direct%20Link-CopyAsPngLikeFigmaWithDialog.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/CopyAsPngLikeFigmaWithDialog)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script to rasterize selected objects at high resolution and copy as bitmap to clipboard, similar to Figma’s “Copy as PNG”.
- Rasterizes at 600dpi, then scales up to the nearest even integer multiple of 72ppi for optimization.

### Main Features

- Automatically creates and manages temporary layers and groups
- Rasterizes at 600dpi, adjusts scaling to an even integer multiple
- Progress bar shows processing status
- Automatically deletes temporary objects and restores original selection
- Supports Japanese and English environments

### Process Flow

1. Select objects
2. Duplicate to temporary layer and group
3. Rasterize at 600dpi
4. Scale up to even integer multiple of 72ppi
5. Copy bitmap to clipboard
6. Delete temporary objects and restore selection

### Update History

- v1.0.0 (20250502): Initial version
- v1.0.1 (20250603): Adjusted scaling to even integer multiples, cleaned up comments