# SmartTableMaker

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartTableMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/SmartTableMaker.jsx)

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

In the height (%) and Width fields, Up/Down steps by ±1, Shift+Up/Down by ±10 (snapping to multiples of 10), and Option+Up/Down by ±0.1.

### Update History

- v1.0 (20260124): Initial version
- v1.1 (20260126): Added Balance (None / Left / Right) and Width so the left/right ratio can be tuned; Width takes the inter-object gap as its maximum and supports slider, numeric input and arrow keys
- v1.2 (20260131): Introduced a PreviewManager based on `app.undo()` so the preview does not pollute the Undo history; on OK the preview is rolled back and the real run happens once, so a single Ctrl+Z reverts it

### Script info

- Version: v1.1
