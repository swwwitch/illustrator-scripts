# SplitBackgroundForTwo

[![Direct](https://img.shields.io/badge/Direct%20Link-SplitBackgroundForTwo.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/SplitBackgroundForTwo.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitBackgroundForTwo.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Select two objects (text, paths, groups, and so on) and the script draws a two-part background behind them.
The Direction setting switches between a left/right split and a top/bottom split.
A dialog asks for a size ratio (%) and shows a preview before you close it (200% by default).
The preview is drawn on a dedicated layer and swapped out, so shapes and strokes are never created twice on OK.

The split is derived from the gap between the two objects.
The Balance panel offers None / Left / Right (None / Top / Bottom in vertical mode):

- None: even split down the middle
- Left (Top): the left (top) margin is set by Width; the other side is calculated
- Right (Bottom): the right (bottom) margin is set by Width; the other side is calculated

Width is set with a slider or a numeric field, and its maximum is taken automatically from the gap between the two selected objects.
To reduce drift caused by side bearings and similar, the text is temporarily converted to outlines to measure its bounding box, and the temporary items are deleted immediately afterwards.
The background rectangles are then placed behind the selected objects and previewed live inside the dialog; the original objects are never modified.

In the size (%) and Width fields, Up/Down steps by ±1, Shift+Up/Down by ±10 (snapping to multiples of 10), and Option+Up/Down by ±0.1.

### Update History

- v1.0 (20260124): Initial version
- v1.1 (20260126): Added Balance (None / Left / Right) and Width so the left/right ratio can be tuned; Width takes the inter-object gap as its maximum and supports slider, numeric input and arrow keys

### Script info

- Version: v2.9
