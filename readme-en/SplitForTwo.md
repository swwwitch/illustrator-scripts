# Replace an object with a two-color split background

[![Direct](https://img.shields.io/badge/Direct%20Link-SplitForTwo.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/SplitForTwo.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitForTwo.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Select a single object (text, a path, a group and so on) and run the script to split its bounding box left/right or top/bottom and replace the object with a two-color background (the original object is deleted).
- The split starts as top/bottom for a tall object and left/right otherwise.
- Set the split position, fills, strokes and corners in the dialog, adjust them while watching the preview, and click OK to commit.

### Main Features

- Split position: Balance sets the width and share (%) of each half with fields and a slider. Turn on Square to split where that side becomes a square.
- Fills: choose whether each half is filled and pick its color. Clicking a swatch opens a color picker with white, black, RGB, CMYK and grayscale. By default the left (top) half is light gray (RGB 220) and the right (bottom) half dark gray (RGB 128).
- Strokes: add an outer frame around the whole shape and a divider along the split, with their weight and color.
- Corners: round each of the four corners on its own, or turn on Pill shape to make both ends semicircles. With Link on, the top-left setting applies to all four corners.
- Accurate measuring: text is measured from an outlined duplicate, so side bearings do not shift the background (the duplicate is deleted right away). Clip groups are measured by their mask.
- Placement: the result goes on the layer that held the original object. Fills go to the back of the layer and the frame and divider to the front; with Group items on they are grouped together, and the group sits at the front of the layer.

### Usage

1. Select one object
2. Run the script
3. Adjust the values while watching the preview, then click OK

### Options

- Split Direction: Left/Right or Top/Bottom. Press H for left/right and V for top/bottom
- Balance: the width (ruler units) and share (%) of the left (top) and right (bottom) halves. Changing any of them updates the rest. The slider sets the same value; hold Option while dragging to snap to whole units, and Shift+Up/Down moves it by 10% of its range
- Square: splits where that side becomes a square
- Fill: turns Left and Right (Top and Bottom for a top/bottom split) on or off and sets their colors
- Stroke: Outer frame and Divider (also toggled with the F and D keys), Stroke width (in the preference's stroke unit) and Color. Stroke width and Color are dimmed when neither the frame nor the divider is in use
- Corner radius: Pill shape, Link, and TL / BL / TR / BR with their radius (ruler units)
- Group items: groups the resulting objects together (on by default)
- In the number fields, Up/Down steps by ±1, Shift+Up/Down by ±10 (snapping to multiples of 10), and Option+Up/Down by ±0.1

### Notes

- Clicking OK deletes the original object. Cancel leaves it as it was and restores the selection.
- While the dialog is open, the original object is hidden and the preview is drawn on a dedicated layer (`__SplitForTwo__PreviewLayer__`) that is removed when the dialog closes.
- Clicking OK with an invalid stroke width while the frame or divider is on returns to the stroke width field.
- Pill shape applies a live pathfinder (Add) to the fills and the frame.
- Fill, stroke and corner settings and the dialog position are kept until Illustrator quits. The split direction and Balance are set afresh for each object, and Group items always starts on.

### Article

- [DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n1b7b8759e53b)

### Update History

- v2.9.2 (20260314): Updated the version string and the update date.
- v2.9.4 (20260922): Fixed tooltips that showed internal names such as "tipColorType". Fixed Cancel still moving the original object to the front and dropping the selection. Fixed the object's layer being switched to Print. Fixed preview layers piling up and staying behind for objects on sublayers; the group is now created on the original object's layer. Fixed the default stroke becoming 0 with inch or cm stroke units, which kept the preview from appearing. Fixed the preview disappearing when the stroke width was blank even though neither the frame nor the divider was in use; OK with an invalid stroke width while the frame or divider is on now returns to that field. Fixed the Balance width and percent fields not accepting a decimal point, and Shift+Up/Down not moving the slider left (up) of center. Fixed values typed into the color picker fields not being applied. Fixed an error on whitespace-only text, and misbehavior when run with one character selected while editing text. Fixed the last values getting mixed with SplitBackgroundForTwo, which used the same #targetengine. Fixed "Bottom:" being cut off in English after switching to a top/bottom split, and field labels now end with a colon. Clip groups are now measured by their mask. Rounding now follows the unit size (1 decimal for pt, 2 for mm, 3 for in, cm and larger units), and the last values are kept in pt so they survive a unit change. The color picker position is now kept until Illustrator quits. The object is measured once when the dialog opens, which speeds up the preview. Internal cleanup
