# SplitForTwo

[![Direct](https://img.shields.io/badge/Direct%20Link-SplitForTwo.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/SplitForTwo.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitForTwo.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Select a single object (text, a path, a group, and so on) and the script splits its bounding box in two — horizontally or vertically — and draws a two-colour background behind it.
The preview is drawn on a dedicated layer and swapped out, so shapes and strokes are never created twice on OK.

- Split method: choose Left/Right or Top/Bottom
- Balance: adjust the width and ratio of the left and right (or top and bottom) halves with numeric fields and sliders
- Stroke: toggle the outer frame and the divider, and set their weight and colour
- Options: apply a corner radius or a pill shape
- Colours are picked from an RGB / CMYK / grayscale colour picker
- Presets can be saved, recalled and exported

To reduce drift caused by side bearings and similar, the text is temporarily converted to outlines to measure its bounding box, and the temporary items are deleted immediately afterwards.
The background rectangles are then placed behind the selected object and previewed live inside the dialog; the original object is never modified.

In the Width field, Up/Down steps by ±1, Shift+Up/Down by ±10 (snapping to multiples of 10), and Option+Up/Down by ±0.1.

### Update History

- v2.9.2 (20260314): Updated the version string and the update date.

### Script info

- Version: v2.9.2
