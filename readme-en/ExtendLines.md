# ExtendLines

[![Direct](https://img.shields.io/badge/Direct%20Link-ExtendLines.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/ExtendLines.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExtendLines.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Takes pairs of adjacent anchor points from the paths in the selection — groups and compound paths included — and draws each as a construction line extended across the drawing area.

### Features

- With Straight on, only straight segments are used; with it off, only curved ones
- Arc to circle: estimates a circle from a Bézier segment. When the segment is not a true arc, the Arc option chooses ignore / chord / extended chord
- Stroke weight follows the "strokeUnits" preference (default equivalent to 0.1 mm)
- With preview on, the lines are drawn on a temporary layer while the dialog is open and removed when it closes

### Usage

1. Select the paths.
2. Run the script.
3. Set the options and run it.

### Notes

- When the selection does not intersect the active artboard, a virtual rectangle four times the selection's width and height is centered on it, and the lines are drawn within that.

### Update History

- v1.0 (2026-02-28)
