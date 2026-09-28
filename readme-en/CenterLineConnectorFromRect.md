# Build rules automatically from pasted rectangles

[![Direct](https://img.shields.io/badge/Direct%20Link-CenterLineConnectorFromRect.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/CenterLineConnectorFromRect.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterLineConnectorFromRect.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Builds rules (a grid or a frame) automatically from rectangles pasted in from Excel or similar
- Depending on how the lines relate to each other it also grids, joins and merges them

### Main Features

- Optional center-lining, rotation correction, exclusion rules, grid mode and outer-frame rectangles
- Stroke weight follows Illustrator's `strokeUnits` and the short-edge threshold follows `rulerType`, with display values and internal point conversion handled in one place
- Finishing passes include unifying stroke weights, picking a representative value or a fixed one, converting to print black, and grouping
- An IIFE structure separates UI construction, event wiring, value reading and the execution flow for maintainability
- Unstable Illustrator DOM operations (move, remove, selection, and so on) go through safe helpers
- Results are built on a dedicated working layer, avoiding dependence on locked layers

### Update History

- v1.0.0 (2025-06-12): Initial release
- v1.6.0 (2026-04-27): Separated the UI structure, reorganized unit handling, added safe-operation helpers, cleaned up naming, and added a dedicated output layer
- v1.6.5 (2026-04-27): Improved the UI structure (added a post-processing panel, reorganized options, improved labels) and refined the output layer design
- v1.7.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.7.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.7.2 (2026-09-28): The button row is now built with the shared part
- v1.7.3 (2026-09-29): Dialog opacity changed to 98%

### Script info

- Version: v1.7.3
