# Build a two-part background behind two objects

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartTableMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/single-function/SmartTableMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartTableMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Select two objects (text, paths, groups, and so on) and the script draws a two-part background behind them, split left and right.
A dialog asks for a height ratio (%) and shows a preview before you close it (200% by default).

The left and right background widths are derived from the gap between the two objects.
The Balance panel offers None / Left / Right:

- None: even split down the middle
- Left: the left margin is set by Width (the right side is calculated)
- Right: the right margin is set by Width (the left side is calculated)

Width is set with a slider or a numeric field, and its maximum is taken automatically from the gap between the two selected objects.
To reduce drift caused by side bearings and similar, the text is temporarily converted to outlines to measure its bounding box, and the temporary items are deleted immediately afterwards.
The background rectangles are then placed behind the selected objects and previewed live inside the dialog; the original objects are never modified.

In the height (%), stroke weight, corner radius and Width fields, the stepper buttons on the left or Up/Down step to the next whole number (1.5 → 2), Shift-click or Shift+Up/Down to the next multiple of 10, and Option-click or Option+Up/Down by ±0.1.

### Update History

- v1.0 (20260124): Initial version
- v1.1 (20260126): Added Balance (None / Left / Right) and Width so the left/right ratio can be tuned; Width takes the inter-object gap as its maximum and supports slider, numeric input and arrow keys
- v1.2 (20260131): Introduced a PreviewManager based on `app.undo()` so the preview does not pollute the Undo history; on OK the preview is rolled back and the real run happens once, so a single Ctrl+Z reverts it
- v1.2.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.2 (20260928): The button row is now built with the shared part
- v1.2.3 (20260929): Dialog opacity changed to 98%
- v1.2.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.2.5 (20260930): Dropped the script's own rightward shift of the dialog on first open. Button rows with only right-side buttons are now centered
- v1.2.6 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.2.7 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part

### Script info

- Version: v1.2.5
