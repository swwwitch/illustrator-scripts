# Build circles, polygons, stars, superellipses and Reuleaux shapes

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartShapeMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/SmartShapeMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartShapeMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script builds circles, regular polygons, stars, superellipses and Reuleaux (constant-width) shapes from a single dialog.

Instead of switching between the Rectangle, Ellipse, Polygon and Star tools, you set the side count, width, rotation, fill and stroke, corner smoothing and anchor options in one place, with a live preview.

### Movie

https://youtu.be/EoUUIdbC0IU

<img alt="Create Basic Shapes dialog" src="../png/ss-1018-1388-144-20260927-045503.png" width="50%" />

## Key features

### Side count

- Pick 0 (circle), 3, 4, 5, 6 or 8 with radio buttons
- Choose "Other" to enter 3 to 36 in a field or with a slider

### Circle options (side count 0)

- Superellipse, with a slider on the same row setting the exponent (1.5 to 6.0) that controls its shape
- Anchor count of 2, 3, 4, 5 or 6 (4 uses Illustrator's ellipse, any other count builds a custom smooth path)
- While the superellipse is on, the anchor count cannot be chosen and the rotation stays off

### Star (side count of 1 or more)

- Turning "Star" on enables the inner radius, entered as a percentage of the outer radius (which equals the width)
- Turning "Pentagram" on sets the side count to 5 and adjusts the angle so the shape reads as five line segments
- The inner radius field and slider are dimmed while "Star" is off

### Width and rotation

- The width is entered in Illustrator's ruler unit
- The rotation angle is typed into a field; the ∧∨ stepper buttons and the arrow keys step it to the next whole number (Shift: next multiple of ten, Option: ±0.1)
- Changing the side count fills the field with the angle that levels the base (360 / sides / 2, or 45 degrees for a circle)
- With a side count of 3, the Triangle panel selects the direction: right, left or down

### Fill and stroke

- Fill and stroke can each be turned on or off, and clicking a swatch opens the color picker
- Stroke width follows Illustrator's stroke unit
- Opacity from 0 to 100 percent, set in a field or with a slider

### Rounded corners (side count 4)

- A corner radius, defaulting to 15 percent of the width
- Smoothing from 0 to 150 percent. At zero the shape uses the Round Corners effect; above zero it becomes a custom bezier path with smoothed corners

### Anchor points

- "Add Anchors (Roughen)" adds anchor points (a value of 1 on a non-circle uses the Add Anchor Points command instead)
- "Split at Anchor Points" breaks the shape into one open path per segment, with a selectable line cap: butt, round or projecting

### Options

- "Convert to Live Shape" converts the result into a live shape after it is confirmed
- "Reuleaux (Constant-Width)" turns each edge of an odd-sided polygon into a circular arc. The amount runs from 0 to 200 percent and resets to 100 whenever the option is enabled

### Other

- Live preview that leaves the undo history clean; the confirmed result is undone in a single step
- "Fit to Window" refits the view to the shape whenever the width changes. A percentage sets how much of the window the shape fills (65 by default); cancelling restores the original view
- The dialog position and settings persist while Illustrator is running
- Localized UI (Japanese / English)

## How to use

1. Run `SmartShapeMaker.jsx` with a document open.
2. Set the side count, width and options in the dialog.
3. Adjust the values while watching the preview.
4. Click OK to create the shape at the center of the document window.

## Keyboard shortcuts

| Key | Action |
| --- | --- |
| E | Circle (side count 0) |
| A | Toggle Rotate |
| S | Toggle Star |
| P | Toggle Pentagram |
| D | Toggle Split at Anchor Points |
| L | Triangle left; also sets the side count to 3 |
| R | Triangle right; also sets the side count to 3 |
| B | Triangle down; also sets the side count to 3 |
| Option (Alt) + 3, 4, 5, 6, 8 | Select that side count (works inside text fields too) |

## Why use this over the standard tools

- No need to switch between four separate tools
- The rotation angle that levels a regular polygon is offered automatically
- Radio buttons instead of spin buttons
- The preview is visible before you press OK, which matters most for the Star tool
- The Pentagram option adjusts the angle so a five-sided star reads as five line segments
- The inner radius of a star is given as a percentage rather than a length (the outer radius equals the width)
- Shapes the standard tools cannot make: superellipses, Reuleaux constant-width shapes and smoothed corners

## Notes

- Values are entered in Illustrator's ruler unit; the stroke width follows the stroke unit.
- Shapes are created at the center of the document window, not the artboard.
- "Convert to Live Shape" is unavailable while any of these is active: Split at Anchor Points, Superellipse, a circle anchor count other than 4, Reuleaux, Add Anchors (Roughen), or corner smoothing.
- Reuleaux applies only to odd-sided polygons (3, 5, 7 and so on) and cannot be combined with a star.
- The rotation is forced off while Pentagram or Superellipse is active.
- "Split at Anchor Points" always opens in the off state.
- The dialog position and settings persist only while Illustrator is running and reset on restart.

## Original / Acknowledgements

- Original idea: Seiji Miyazawa (Sankai Lab)
- Corner smoothing: Shingo Kurono, [Reproducing corner smoothing in Illustrator (Japanese)](https://note.com/shingokurono/n/n348a3e73a465)

## Article

[Create squares, circles and triangles with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n005a7087f9c3)

## Changelog

- v2.4.5 (2026-09-30): Dropped the script's own rightward shift of the dialog on first open
- v2.4.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v2.4.3 (2026-09-29): Dialog opacity changed to 98%
- v2.4.2 (2026-09-28): The button row is now built with the shared part. Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held)
- v2.4.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v2.4.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v2.3.0 (2026-09-27): Removed the View Zoom slider. "Fit to Window" gained a percentage field, 65 by default, and Cancel now restores the view position as well. Added Option (Alt) + 3/4/5/6/8 shortcuts for the side count. The "4 (Square)" label, colons after "Fill:" and "Stroke:", a percent sign on the smoothing value, the superellipse exponent reduced to a slider on the checkbox row, and reworded labels such as "Anchor Count", "Rounded Corners", "Convert to Live Shape", "Fit to Window" and "Cap". Added tooltips and consolidated duplicated internal code
- v2.2.4 (2026-09-27): Fixed number fields that would not accept a decimal point (the value was rewritten on every key, not just the arrow keys). Units now come from the shared unit table (adds H, ft, yd and m; inches read "in"). Renamed the English dialog title to "Create Basic Shapes". Added tooltips to the controls and unified how field labels get their colon
- v2.2.3 (2026-09-21): Fixed the dialog title showing ColorPicker.jsx's version instead of this script's. Fixed the color picker not opening when a color swatch is clicked (ColorPicker.jsx v1.0.2)
- v2.2.0 (2026-07-31): Added the basic info block and JSDoc, wrapped the script in an IIFE, renamed variables and panels, and fixed when the parameters are captured on OK
- v1.0.0 (2025-05-02): Initial version
