# List and edit gradient stops

[![Direct](https://img.shields.io/badge/Direct%20Link-EditGradientStops.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/EditGradientStops.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditGradientStops.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Lists the gradient stops used by the selected objects and lets you change their color, location and opacity, or swap stops.
It can also turn a gradient into a segmented one (hard edges, no blending) whose colors switch at the percentages you enter, or build one from the colors of objects without gradients.

### Features

- Collects fill and stroke gradients, including those inside groups, compound paths and text, and lists them in one panel per gradient
- Each stop is one row: color chip, color values (CMYK, RGB, gray, Lab, or swatch name and tint), location and opacity, in location order
- Click a color chip to change the color with Illustrator's standard color picker
- Change location and opacity with the ∧∨ buttons, the arrow keys or by typing (arithmetic allowed), from 0 to 100%
- The ↑↓ buttons at the end of a row swap color and opacity with the stop above or below (locations stay)
- The duplicate button at the end of a row adds a copy of the stop right below it (placed halfway to the next stop, or to the previous stop on the last row)
- The delete button at the end of a row removes the stop (at least two stops remain)
- Changes show on the objects right away (preview is always on)
- Gradients with identical stops share one panel, and changes go to every fill and stroke that uses them
- **Segment**
  - Segment... at the bottom left opens a dialog box for the boundary positions in % and makes the listed gradients segmented (Boundary 1, Boundary 2, ..., one field fewer than the colors)
  - When the listed gradients have different numbers of colors, only those with the same count as the first gradient with two or more colors are changed
  - The boundaries start at the current positions for a segmented gradient, or evenly spaced for a smooth one
  - The positions are applied as you type
  - OK replaces the panel's rows with the segmented stops, which you can keep editing (they are written when you click OK in the main dialog box)
- **Round Dashes** (at the bottom left of the Segment dialog box, shown only when a target stroke uses a gradient)
  - Sets dashes (1 pt dash, blank gap), round caps and Align Dashes to Corners and Path Ends on gradient strokes (stroke weight, join and miter limit stay)
  - Applied as soon as you click it; Cancel restores the dashes and caps (in either the Segment dialog box or the main one)
- **Create**
  - Run the script with two or more differently colored objects without gradients selected, enter the boundary positions, and a segmented gradient is created from those colors
  - Colors are taken in layout order (left to right for a horizontal row, top to bottom for a vertical one)
  - A rectangle filled with the new gradient is placed below the selection (horizontal) or to its right (vertical), one rectangle apart (angle 90° when vertical)
  - Boundaries are remembered per color count and used as the next defaults (until Illustrator quits)

### Usage

1. Select objects that use gradients.
2. Run the script.
3. Change the color, location and opacity, then click OK. Cancel restores the original stops.

To create a new segmented gradient, select two or more differently colored objects without gradients, run the script, enter the boundary positions and click OK.

### Notes

- Gradients in fills or strokes added in the Appearance panel are not covered.
- For text, the fill and stroke returned for the whole text frame's character attributes are read and written.
- The color picker returns colors in the document's color mode, so picking a new color for a swatch color turns it into a process color.
- Up to 10 gradients are listed (change MAX_LISTED_GRADIENTS at the top of the script).
- Segment counts consecutive stops with the same color and opacity as one color. The two colors on each side of a boundary sit at ±0.01% of it (change BOUNDARY_OFFSET at the top of the script).
- Round Dashes runs a stroke action, so arrowheads are removed and the stroke is aligned to center. Align Dashes to Corners and Path Ends stays on even after Cancel. Strokes on text are not covered. Change the dash values with STROKE_DASH_PATTERN at the top of the script.
- Segment handles up to 10 colors (change MAX_SEPARATE_COLORS at the top of the script).

### Article

[List and edit gradient stops (Japanese)](https://note.com/dtp_tranist/n/n524fafa198dd)

### Update History

- v1.1.0 (2026-11-20) Merged in the features of AdjustSeparateGradient. Segment... at the bottom left turns the gradients into segmented ones at the boundary positions you enter. Removed the Preview checkbox; the preview is now always on. When a stroke uses a gradient, Round Dashes in the Segment dialog box sets dashes (1 pt dash, blank gap), round caps and Align Dashes to Corners and Path Ends. With two or more differently colored objects without gradients selected, creates a new segmented gradient from their colors
- v1.0.1 (2026-11-20) Clicking the script name and version at the bottom left opens About This Script (copy the script name, open the README and other links, show in Finder; shared part updated)
- v1.0.0 (2026-11-20) First release
