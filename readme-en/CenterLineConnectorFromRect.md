# CenterLineConnectorFromRect

[![Direct](https://img.shields.io/badge/Direct%20Link-CenterLineConnectorFromRect.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/CenterLineConnectorFromRect.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterLineConnectorFromRect.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Updated: 2026-04-27
- Builds rules (a grid or a frame) automatically from rectangles pasted in from Excel or similar
- Depending on how the lines relate to each other it also grids, joins and merges them

### Main Features

- Optional center-lining, rotation correction, exclusion rules, grid mode and outer-frame rectangles
- Stroke weight follows Illustrator's `strokeUnits` and the short-edge threshold follows `rulerType`, with display values and internal point conversion handled in one place
- Finishing passes include unifying stroke weights, picking a representative value or a fixed one, converting to print black, and grouping
- An IIFE structure separates UI construction, event wiring, value reading and the execution flow for maintainability
- Unstable Illustrator DOM operations (move, remove, selection, and so on) go through safe helpers
- Results are built on a dedicated working layer, avoiding dependence on locked layers

### Script info

- Version: v1.6.5
