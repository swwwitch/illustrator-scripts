# Turn strokes into two-color dashed graphs

[![Direct](https://img.shields.io/badge/Direct%20Link-StrokeGraph.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/StrokeGraph.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/StrokeGraph.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Turns the strokes of the selected paths into round dashes split into a base color and an accent color, making stroke graphs such as ring charts and progress bars.
Works without an existing gradient stroke, and each path gets its own gradient, so duplicated objects elsewhere stay unchanged.

### Features

- Sets dashes (1 pt dash, blank gap), round caps and Align Dashes to Corners and Path Ends on the strokes of the selected paths (including those in groups and compound paths)
- Reverses the path direction and rotates the paths by 90° (puts the start of a circle at the top)
- Turns the stroke into a two-color segmented gradient: base from the start, accent over the last accent % (20% puts the boundary at 80%)
- Set in the dialog box
  - Base and accent: click a chip to change the color with Illustrator's standard color picker
  - Accent %: the ∧∨ buttons, the arrow keys or typing (arithmetic allowed)
  - Weight: shown in the stroke unit preference
- Changes show right away (preview is always on)
- Each path gets its own gradient (Stroke Graph, Stroke Graph 1, ...), so unselected objects sharing the same gradient stay unchanged
- Defaults
  - For a gradient stroke, the start color becomes the base and the end color the accent; if it is already split in two, the accent % comes from the current boundary
  - Otherwise, the stroke color (light gray when there is no stroke) becomes the base, red the accent, and the accent % is 20%
- Cancel restores the stroke settings, direction, rotation and gradients

### Usage

1. Select paths.
2. Run the script.
3. Set the colors, accent % and weight, then click OK.

### Notes

- Scripts cannot set the stroke gradient type. A gradient newly added to a stroke without one uses Apply Gradient Within Stroke (the default), so switch to Apply Gradient Along Stroke in the Gradient panel for a ring chart. Strokes that already had a gradient keep their type.
- The dashes are set with a stroke action, so arrowheads are removed and the stroke is aligned to center. Align Dashes to Corners and Path Ends stays on even after Cancel.
- Each run adds one gradient swatch per path.
- Running it twice on the same path reverses and rotates the path again.
- Change the rotation (GRAPH_ROTATION), the default accent % (DEFAULT_ACCENT_PERCENT), the default colors (DEFAULT_BASE_COLOR, DEFAULT_ACCENT_COLOR) and the dash values (STROKE_DASH_PATTERN) at the top of the script.

### Update History

- v1.0.0 (2026-11-20) First release
