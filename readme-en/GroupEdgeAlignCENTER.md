# GroupEdgeAlignCENTER

[![Direct](https://img.shields.io/badge/Direct%20Link-GroupEdgeAlignCENTER.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/GroupEdgeAlign-7/GroupEdgeAlignCENTER.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlignCENTER.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Takes the edges or the center of the selected objects and aligns them to the center.

The target is the edge of the active artboard, or a matching guide. The direction is taken from the filename.

### Features

- When `USE_GUIDES` is true, a guide for the chosen direction is looked up and snapped to according to `GUIDE_SEARCH_MODE`
- Falls back to the artboard edge when no matching guide exists
- When `USE_GUIDES` is false, it always aligns to the artboard edge

### Usage

1. Select the objects to align.
2. Run the script.

### Notes

- The center options never use guides; they always align to the center of the artboard.
- Use GroupEdgeAlign.jsx to pick the direction from a dialog.

### Update History

- v1.0 (2025-04-06)
