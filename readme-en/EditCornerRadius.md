# Edit the corner radius of rounded rectangles

[![Direct](https://img.shields.io/badge/Direct%20Link-EditCornerRadius.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/EditCornerRadius.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditCornerRadius.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Edits the corner radius of rounded rectangles in a dialog
- Choose the target: "Selected objects only", "Current artboard only" or "Entire document"
- The dialog shows the measured current radius
- OK rebuilds each path with the corners set to that radius

### Main Features

- Values use the ruler units
- The target starts at "Selected objects only" when something is selected, otherwise at "Current artboard only" ("Selected objects only" is unavailable without a selection)
- "Current artboard only" and "Entire document" include rectangles inside groups. A rectangle counts as on the artboard when it overlaps it
- "Keep zero radii at zero" (on by default): corners without rounding stay square and only rounded corners get the new radius. Turn it off to set all four corners to the radius
- Up/Down arrow keys change the value (Shift: 10, Option: 0.1)
- "Preview" shows the result while you adjust
- Multiple selections are supported (the same radii are applied to every rectangle)
- To round a plain rectangle, turn off "Keep zero radii at zero"
- The direction of the anchor points (clockwise / counterclockwise) is kept as in the original path
- "Include the Round Corners effect" (off by default): rectangles rounded by Effect > Stylize > Round Corners are included, and the effect is reapplied with the new radius. The rectangle's appearance is cleared first, so its other effects and extra fills or strokes are removed (the basic fill, stroke and opacity are restored)
- "Convert to Round Corners effect" (off by default): rectangles with all four corners rounded are made square and rounded with the Round Corners effect (other appearance attributes are kept). Unavailable while "Keep zero radii at zero" is on

### Usage

1. Select the rectangles whose corners you want to change (not needed for the artboard or document targets)
2. Run the script
3. Choose the target, enter the radius and click OK

### Notes

- Only rectangles (rounded or not) aligned to the horizontal and vertical axes are changed. Rotated rectangles, other paths, locked or hidden objects, guides and parts of compound paths are left as they are
- With "Selected objects only", select rectangles inside groups directly, for example with the Direct Selection tool
- With "Selected objects only", the dialog shows how many selected objects are skipped
- OK is unavailable when the target has no rectangles
- The initial radius is the average of the rounded corners (zero radii excluded) in the target chosen when the dialog opens
- Values over half the shorter side are limited to half of it
- With "Include the Round Corners effect" on, each rectangle without rounded path corners is duplicated and expanded to measure it, so many targets take time

### Update History

- v1.1.0 (2026-09-26): Merged the radius input into one field so the corners get the same radius; added "Keep zero radii at zero", "Include the Round Corners effect", "Convert to Round Corners effect" and a target switch (selection / current artboard / entire document); the initial radius is now the average of the targets
- v1.0.0 (2026-09-26): Initial release

### Script info

- Version: v1.1.0
