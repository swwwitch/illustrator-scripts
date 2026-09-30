# Rebuild horizontal and vertical lines into a grid

[![Direct](https://img.shields.io/badge/Direct%20Link-RectangularGridReverseTool.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/RectangularGridReverseTool.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangularGridReverseTool.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Analyzes the selected horizontal and vertical lines and reconstructs them into a rectangular grid.

Uneven rules and layouts containing merged cells are tidied into a regular lattice.

### Features

- Pre-processing: split the outer frame into four sides
- Distribution: none / forced even / even with merged-cell support
- Scope: control which of the vertical and horizontal rules are equalized
- Stroke (post-process): projecting caps, dashed to solid, stroke weight (max / min / average / explicit)
- Post-processing: convert the outer frame to a rectangle, group the result, center point text vertically in cells
- Preview: check the result without closing the dialog

### Usage

1. Select the horizontal and vertical lines.
2. Run the script.
3. Set pre-processing, distribution, scope, stroke and post-processing, and confirm while watching the preview.

### Update History

- v1.3.5 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.3.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.3.3 (20260929): Dialog opacity changed to 98%
- v1.3.2 (20260928): The button row is now built with the shared part
- v1.3.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.0
