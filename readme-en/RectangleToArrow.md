# Turn rectangles into arrows

[![Direct](https://img.shields.io/badge/Direct%20Link-RectangleToArrow.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/RectangleToArrow.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArrow.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Turns each selected rectangle into an arrow that fills its bounds.
- Choose Fill, Stroke (Butt Cap) or Stroke (Round Cap), and adjust the thickness with a live preview.

<img alt="The Rectangle to Arrow dialog" src="../png/ss-430-518-144-20261010-165317-s.png" width="50%" />

### Features

#### Type

- Pick the type with three stacked buttons that show the arrow shape. The selected one has a gray background
  - Fill: a single closed path. The original path is rewritten, so its effects and stacking order stay
  - Stroke (Butt Cap): a straight shaft and a chevron head stroked with butt caps and miter joins
  - Stroke (Round Cap): the same with round caps and round joins
- The two stroke types apply the Outline Stroke effect to each line, group the lines, and apply the Pathfinder (Add) effect to the group. They stay live, so you can still edit the stroke weight and so on
- Direction: wide rectangles point right, tall ones point up
  - Option-click (Alt-click) a button to reverse the direction (left or down). Do it again to switch back
  - Command-Option-click (Ctrl-Alt-click) a button to toggle a double-headed arrow
  - Direction and double-headed apply to all three types. The button drawings follow the current direction

#### Thickness

- A percentage of the rectangle's short side (1–100%): the shaft thickness for Fill, the stroke weight for the stroke types
- Starts at 45% for Fill and 15% for the stroke types. Each type keeps its own value when you switch
- Use the ∧∨ buttons or the ↑↓ keys (Shift for the next multiple of 10, Option for 0.1 steps)

#### Shape and color

- The head is right-angled and half the short side long. On short rectangles it is capped at the full length (single) or half the length (double-headed)
- Stroked arrows fit the path to the rectangle, so the stroke weight sticks out past its edges
- The color comes from the rectangle's fill (its stroke when unfilled, black when neither)

#### Preview

- The preview is always on, from the moment the dialog opens. Copies are turned into arrows while the original rectangles are hidden; Cancel restores them

### How to use

1. Select the rectangles to turn into arrows (rectangles inside groups count; several at once is fine)
2. Run the script and set the type and thickness in the dialog
3. Click OK to replace them

### Notes

- Only unrotated rectangles (four points, straight horizontal and vertical sides) are converted. Rotated or rounded rectangles and locked or hidden items are skipped
- Stroked arrows do not keep effects that were on the original rectangle
- Stroked arrows apply their effects through menu commands on every preview update, so redrawing can slow down with many rectangles

### Article

https://note.com/dtp_tranist/n/n789072361c12

---

### Update History

- v1.0.0 (2026-10-10) Initial release

### Script info

- Version: v1.0.0
- First release: 2026-10-10
- Last updated: 2026-10-10
