# Edit the corner radius of rounded rectangles

[![Direct](https://img.shields.io/badge/Direct%20Link-EditCornerRadius.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/EditCornerRadius.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditCornerRadius.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Edits the corner radius of rounded rectangles in a dialog
- Choose the target: "Selected objects only", "Current artboard only" or "Entire document"
- The dialog shows the measured current radius, and OK rebuilds each path with the corners set to that radius
- Callout shapes and rectangles inside groups, compound paths and compound shapes are included

### Main Features

- Values use the ruler units
- The target starts at "Selected objects only" when something is selected, otherwise at "Current artboard only" ("Selected objects only" is unavailable without a selection)
- With "Current artboard only", a rectangle counts when it overlaps the artboard
- The stepper buttons left of the field and the Up/Down arrow keys step to the next whole number (1.5 → 2; Shift: next multiple of 10, Option: 0.1)
- "Preview" shows the result while you adjust
- Multiple selections are supported (the same radius is applied to every rectangle)
- The direction of the anchor points (clockwise / counterclockwise) is kept as in the original path
- Shapes that are changed
  - Rectangles (rounded or not) aligned to the horizontal and vertical axes
  - Rectangles with a callout tail or similar on a side. The tail stays as it is and only the four corners change (the radius is limited to the base of the tail)
  - Rectangles inside groups, including subgroups (with "Selected objects only", select the group)
  - Rectangles inside compound paths (select the compound path)
  - Rectangles inside compound shapes (see below)
- Compound shapes
  - Compound shapes are converted to a group with Effect > Pathfinder > Add, and the rectangles inside are changed (with any target: selection, artboard or document)
  - Converting turns every shape mode (such as Minus Front) into Add. Name, opacity and blending mode are carried over
  - Compound shapes without rectangles are not converted
  - When a rectangle inside has the Round Corners effect, the effect is removed and the same radius is built into the path (removal uses Clear Appearance, so other effects on that path are removed too; fill, stroke and opacity are kept)
  - When rectangles inside are selected directly with the Direct Selection tool, those paths are edited in place without conversion (shape modes stay)

### Usage

1. Select the rectangles to change, or groups, compound paths or compound shapes that contain them (not needed for the artboard or document targets)
2. Run the script
3. Choose the target, enter the radius and click OK

### Options

- "Keep zero radii at zero" (on by default): corners without rounding stay square and only rounded corners get the new radius. Turn it off to set all four corners to the radius (turn it off to round a plain rectangle)
- "Include the Round Corners effect" (off by default): rectangles rounded by Effect > Stylize > Round Corners are included, and the effect is reapplied with the new radius. The rectangle's appearance is cleared first, so its other effects and extra fills or strokes are removed (the basic fill, stroke and opacity are restored)
- "Convert to Round Corners effect" (off by default): rectangles with all four corners rounded are made square and rounded with the Round Corners effect (other appearance attributes are kept). Unavailable while "Keep zero radii at zero" is on
- The two effect options do not apply to callout shapes or to rectangles inside compound paths and compound shapes

### Notes

- Rotated rectangles, other paths, locked or hidden objects, and guides are left as they are
- A converted compound shape cannot be turned back into a compound shape (Undo still works)
- With "Selected objects only", the dialog shows how many selected objects are skipped. When nothing selected can be changed, it shows "Nothing selected can be changed"
- OK is unavailable when the target has no rectangles
- The initial radius is the average of the rounded corners (zero radii excluded) in the target chosen when the dialog opens
- Values over half the shorter side are limited to half of it (for callout shapes, also to the distance to the base of the tail)
- With "Include the Round Corners effect" on, and when converting compound shapes, items are duplicated and expanded to measure them, so many targets take time

### Update History

- v1.5.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.5.0 (2026-09-28): With "Selected objects only", rectangles inside selected groups are included. When converting a compound shape, the Round Corners effect on released rectangles is removed and the same radius is built into the path
- v1.4.0 (2026-09-28): Compound shapes are converted to a group with the Pathfinder Add effect so their rectangles can be changed (selection, artboard and document)
- v1.3.0 (2026-09-28): Added support for callout shapes (rectangles with a tail on a side), rectangles inside compound paths, and directly selected rectangles inside compound shapes. Fixed the dialog widening from the skipped-objects text when nothing selected was a target
- v1.2.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.1.0 (2026-09-26): Merged the radius input into one field so the corners get the same radius; added "Keep zero radii at zero", "Include the Round Corners effect", "Convert to Round Corners effect" and a target switch (selection / current artboard / entire document); the initial radius is now the average of the targets
- v1.0.0 (2026-09-26): Initial release

### Script info

- Version: v1.5.1
