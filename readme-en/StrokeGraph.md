# Turn strokes into two-color dashed graphs

[![Direct](https://img.shields.io/badge/Direct%20Link-StrokeGraph.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/StrokeGraph.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/StrokeGraph.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Turns the strokes of the selected paths into round dashes split into a base color and an accent color, making stroke graphs such as ring charts and progress bars.
The path's own color becomes the accent, and strokes without a gradient get Apply Gradient Along Stroke from an embedded graphic style.

### Features

- Turns the selected paths (including those in groups and compound paths) into stroke graphs
  - Removes the fill
  - Paths without a gradient stroke get the "stroke-gradient" graphic style embedded in the script, which makes the stroke Apply Gradient Along Stroke (the opacity and blending mode stay as they are). The embedded file is opened only when the document does not have the style yet (no external file needed)
  - Turns the stroke into a two-color segmented gradient: base from the start, accent over the last accent % (20% puts the boundary at 80%)
  - Sets dashes (1 pt dash, blank gap), round caps and Align Dashes to Corners and Path Ends
  - Reverses the path direction (when Reverse Direction is on)
  - Sets the angle of live ellipses to 90° (the angle in Ellipse Properties of the Transform panel; ellipses already at 90° are left as they are)
- Each path gets its own gradient (Stroke Graph, Stroke Graph 1, ...), so unselected objects sharing the same gradient stay unchanged
- Set in the dialog box
  - Base and accent: click a chip to change the color with Illustrator's standard color picker
  - Accent %: the ∧∨ buttons, the arrow keys or typing (arithmetic allowed)
  - Weight: shown in the stroke unit preference
  - Reverse Direction: whether to reverse the path direction (on by default; toggling it reverses right away, swapping where the base and accent go)
- Changes show right away (preview is always on)
- Defaults
  - Without a gradient stroke, the path color (the stroke color, or the fill color when there is no stroke; red when there is neither) becomes the accent, K10 the base, the accent % is 20% and the weight is 10 pt
  - With a gradient stroke (such as a path already made into a graph), the start color becomes the base, the end color the accent and the current weight the weight; if it is already split in two, the accent % comes from the current boundary
- Cancel restores the fills, stroke settings, gradients, direction and ellipse angles

### Usage

1. Select paths such as circles or ellipses.
2. Run the script.
3. Set the colors, accent %, weight and Reverse Direction, then click OK.

### Notes

- Scripts cannot set the stroke gradient type, so Apply Gradient Along Stroke comes from a graphic style. A document the style was imported into keeps the "stroke-gradient" style and its gradient swatch. Strokes that already had a gradient keep their type.
- Applying the style restores the opacity and blending mode, but removes extra strokes and effects in the appearance (Cancel does not bring them back).
- The dashes are set with a stroke action, so arrowheads are removed and the stroke is aligned to center. Align Dashes to Corners and Path Ends stays on even after Cancel.
- Each run adds one gradient swatch per path.
- The ellipse angle applies only to live ellipses. Cancel sets the ellipses whose angle was changed to 0° (those already at 90° stay as they are).
- Running it twice on the same path reverses the path again. Turn off Reverse Direction to keep the current direction.
- Change the ellipse angles (ELLIPSE_ANGLE, ELLIPSE_ANGLE_ON_CANCEL), the default accent % (DEFAULT_ACCENT_PERCENT), the default colors (DEFAULT_BASE_COLOR, DEFAULT_ACCENT_COLOR), the default weight (DEFAULT_STROKE_WIDTH) and the dash values (STROKE_DASH_PATTERN) at the top of the script.

### Update History

- v1.1.0 (2026-11-21) Strokes without a gradient now get Apply Gradient Along Stroke from an embedded graphic style. The path color is now the default accent and K10 the default base, and the default weight is 10 pt. Fills are removed. Ellipses already at 90° keep their angle. Added Reverse Direction to the dialog box and removed the description line at the top. The two stops at the boundary are now at the same location, so no in-between color shows
- v1.0.1 (2026-11-20) Replaced the 90° rotation with setting the angle of live ellipses to 90° (dynamic action)
- v1.0.0 (2026-11-20) First release
