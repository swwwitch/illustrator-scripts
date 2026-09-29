# A persistent palette that scales the selection

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartScalePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/SmartScalePalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartScalePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A persistent palette that scales the selected objects. Keep it open, change the selection and Apply as often as you like.
- Specify a percentage or the finished width and height.

### Main Features

#### Size

- Scale (%): stepper buttons left of the field and the arrow keys step the value
  - ↑↓ = to the next whole number (1.5 → 2)
  - Shift+↑↓ = to the next multiple of 10 (232 → 240)
  - Option(Alt)+↑↓ = ±0.1
- Width / Height: shows the current size of the whole selection (ruler units); entering a target size works out the scale
  - The current size is measured again when you return to the palette and right before Apply
  - When a width or height was entered last, the scale is recomputed from the new measurement to reach that size
  - With the link toggle on the right the aspect ratio is kept; when off, width and height scale separately (the scale field stays empty while they differ)
- Buttons in the right column: 50%, 100% and 200% set the value; +10%, -10%, +1% and -1% add to the current scale

#### Target, Anchor and Options

- Target: Each Object, or As a Whole (the whole selection as one). "As a Whole" never groups the objects, so stacking order and hierarchy stay intact
- Reference Point: any of the nine reference points. The bounds follow the "Use Preview Bounds" preference: including strokes and effects when on, path geometry when off
- Options: whether Corners, Strokes & Effects, Patterns and Gradients scale (corners and strokes & effects switch the Illustrator preferences only while scaling and restore them afterwards)

#### Buttons

- Reset: returns the scale to 100% and every setting to its default
- Close: closes the palette
- Apply: scales the selected objects

### Process Flow

1. Run the script to open the palette
2. Select objects and set the scale (or width and height), target, reference point and options
3. Apply scales them (undo with Illustrator's Undo)

### Notes

- After editing the script, close the open palette before running it again (otherwise the old code keeps running)

### Update History

- v1.0 (20250831) : Initial version
- v1.1.0 (20260929) : Added to the repository. Fixed stroke widths getting thinner even when enlarging. "As a group" no longer groups temporarily, so stacking order and hierarchy stay intact. Added stepper buttons to the number field. "Stroke Width" renamed "Strokes & Effects" and now scales effects too. Added "Corners". Added scale buttons (50%, 100%, 200%, ±10%, ±1%) in a right column. Buttons moved into one row at the bottom. Added a Size panel holding the scale plus width and height fields with a link toggle. Removed "Use preview bounds"; the preference is followed instead. The Target, Reference Point and Options panels sit side by side in three columns. The target choices were renamed Each Object / As a Whole, the anchor panel was renamed Reference Point, and the scale field shows %. Added tooltips. The anchor now offers all nine reference points instead of top left or center. Turned into a persistent palette and renamed from SmartScale.jsx to SmartScalePalette.jsx (the preview was dropped; Apply runs the scaling). The palette reopens where it was last closed. Title changed to "Smart Scale"

### Script info

- Version: v1.1.0
