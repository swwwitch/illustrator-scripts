# Create and adjust segmented gradients at set percentages

[![Direct](https://img.shields.io/badge/Direct%20Link-AdjustSeparateGradient.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/single-function/AdjustSeparateGradient.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustSeparateGradient.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Builds a segmented gradient (hard-edged, with no blending between colors) whose colors switch at the percentages you enter.
With two or more differently colored objects selected it creates a new gradient; with objects whose fill or stroke is a gradient selected it moves their boundaries.

The same features are also in [List and edit gradient stops](EditGradientStops.md) (EditGradientStops).

### Features

- The selection decides what the script does
  - If any object has a gradient fill or stroke, that gradient is adjusted
  - Otherwise a new gradient is built from the colors of the selected objects
- Enter each boundary as a percentage ("Boundary 1", "Boundary 2", …, one field fewer than the colors)
- Change positions with the ∧∨ buttons, the arrow keys or by typing (arithmetic allowed)
- **Create**
  - Colors are taken in layout order (left to right for a row, top to bottom for a column)
  - A rectangle filled with the new gradient is placed one rectangle's width away: below a horizontal selection, or to the right of a vertical one (at a 90° angle)
  - Boundaries are remembered per color count and used as the next defaults (until Illustrator quits)
- **Adjust**
  - Collects fill and stroke gradients, including those inside groups and compound paths
  - Smooth gradients are converted to segmented ones; segmented gradients get their boundaries moved
  - The fields start at the current boundaries for segmented gradients, or at even intervals for smooth ones
  - Preview (on by default) applies the positions as you edit them; Cancel restores the original

### Usage

1. Select two or more objects with different colors, or objects whose fill or stroke is a gradient.
2. Run the script.
3. Enter the boundary positions and click OK.

### Notes

- Adjusting rewrites the gradient swatch itself, so other objects that use the same gradient change too.
- When adjusting, consecutive stops with the same color and opacity count as one color.
- When gradients with different numbers of colors are selected together, only those with the same count as the first one found are changed.
- The two colors at a boundary sit ±0.01% from it (change BOUNDARY_OFFSET at the top of the script).
- Up to 10 colors are supported (change MAX_COLORS at the top of the script).
- Gradients on text, and fills or strokes added in the Appearance panel, are not covered.

### Update History

- v1.0.0 (2026-11-20) First release
