# Split a dashed line into separate segments

[![Direct](https://img.shields.io/badge/Direct%20Link-DashNipper.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/DashNipper.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashNipper.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- With a dashed line selected, the script outlines it, splits it into its dashes, and replaces each dash with a line that looks the same (its center line). The result looks like the original dashed line, but every dash is a separate path.
- With one already-outlined shape selected (a closed path with four anchor points), the script replaces that shape with its center line.

### Main Features

- Caps: dashes from a butt-cap dashed line get butt caps; dashes from a round-cap dashed line (pill shapes with semicircular ends) get round caps.
- Stroke width and color: the stroke width is taken from the original dashed line, and the stroke color from the fill of the outlined dash.
- Curves: dashes on a curved path become center lines that follow the curve. A dash that crosses an anchor point or a mitered corner becomes a center line with an anchor point in the middle.
- Direction: the direction is chosen so that the thickness matches the original stroke width, so dashes shorter than the stroke width are not turned sideways.
- Stacking order: each center line is placed directly in front of its dash (same layer or group), and all new lines are selected afterward.
- Several dashed lines can be processed at once; the stroke width is read from each one.
- Single shapes: rectangles, bands with curved long sides, and rotated rectangles become a center line between the midpoints of the short sides. The stroke width is the short side length.

### Usage

1. Select one or more dashed lines
2. Run the script

To convert an already-outlined shape, select just that one shape and run the script.

### Notes

- Only paths with Dashed Line set in the Stroke panel are processed. Dashes made with a brush or with a second stroke in the Appearance panel are not.
- Selecting a group does not include the dashed lines inside it. Select the dashed lines themselves.
- Other objects selected together with the dashed lines are left as they are.
- The following dashes cannot become center lines and stay outlined (an alert reports how many):
  - Zero-length dashes with round caps (the dots of a dotted line)
  - Dashes that cross a beveled or rounded corner
  - Dashes whose two sides have different numbers of anchor points
  - Dashes whose thickness changes with a width profile
- Dashes with projecting caps become lines with butt caps (the visible length stays the same).
- A single shape with equal thickness in both directions (such as a square) becomes a vertical line.

### Update History

- v1.0.0 (20260922) : Initial release
