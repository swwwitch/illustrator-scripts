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

- v1.2.0
