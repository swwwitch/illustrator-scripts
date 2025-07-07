# SmartArtboardRenamer.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartArtboardRenamer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/SmartArtboardRenamer.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script to batch rename artboards in Illustrator with flexible custom rules.
- Allows combining prefix, suffix, and reference text for advanced naming.

### Main Features

- Prefix/suffix can include sequential numbers (1, 01, A, a) and file name (#FN)
- "Frontmost text" mode uses the topmost text frame per artboard
- "Specify layer" mode combines text from a chosen layer
- Automatically appends "_2", "_3", etc. to avoid duplicate names
- Ignores hidden layers
- Supports specifying target artboards by range

### Process Flow

1. Configure mode, prefix, suffix, and other settings in the dialog
2. Select text reference method and artboard range
3. Rename artboards based on settings
4. Automatically adjust duplicate names if needed

### Update History

- v1.0.0 (20250509): Initial version
- v1.0.1 (20250512): Improved layer reference and UI adjustments