# Turn rectangles into arrows

[![Direct](https://img.shields.io/badge/Direct%20Link-RectangleToArrow.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/RectangleToArrow.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArrow.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Turns each selected rectangle or horizontal rule into an arrow that fills its bounds.
- Choose Fill, Stroke (Butt Cap) or Stroke (Round Cap), and adjust the thickness and height with a live preview.
- Reverse the direction or make a double-headed arrow with the icons, or by modifier-clicking a type button.

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
  - The link (double-headed) and ⇄ (reverse) icons under the type buttons switch them too. Reverse is unavailable for double-headed arrows
  - Direction and double-headed apply to all three types. The button drawings follow the current direction

#### Thickness

- Fill: the shaft thickness as a percentage of the rectangle's short side (1–100%, starting at 45%)
- Stroke types: the stroke weight in the Stroke units set in Preferences. It starts at 15% of the first target's short side, converted to those units
- Switching types also switches the field between % and the stroke units. Each type keeps its own value when you switch
- Use the ∧∨ buttons or the ↑↓ keys (Shift for the next multiple of 10, Option for 0.1 steps)

#### Height

- The arrow height (head width) as a percentage of the rectangle's short side (1–500%, starting at 100%). Shared by all three types
- The shaft thickness and stroke weight do not change. On filled arrows, the shaft is capped at the height when the height is smaller
- The head length follows the height, keeping the head right-angled

#### Horizontal rules

- Besides rectangles, straight horizontal open paths (two points) are converted too. Since their height is 0, a quarter of the line's length is used as the short side

#### Shape and color

- The head is right-angled and half the height long. On short rectangles it is capped at the full length (single) or half the length (double-headed)
- Stroked arrows fit the path to the rectangle, so the stroke weight sticks out past its edges
- The color comes from the rectangle's fill. When the fill is none or white, the stroke color is used; with neither fill nor stroke, black

#### Preview

- The preview is always on, from the moment the dialog opens. Copies are turned into arrows while the original rectangles are hidden; Cancel restores them

### How to use

1. Select the rectangles or rules to turn into arrows (items inside groups count; several at once is fine)
2. Run the script and set the type, thickness and height in the dialog
3. Click OK to replace them

### Notes

- Only unrotated rectangles (four points, straight horizontal and vertical sides) and horizontal rules (two-point open paths without handles) are converted. Rotated or rounded rectangles, slanted or vertical lines, and locked or hidden items are skipped
- Stroked arrows do not keep effects that were on the original rectangle
- Stroked arrows use the same stroke weight for every rectangle (the filled shaft is a percentage of each rectangle's short side)
- Stroked arrows apply their effects through menu commands on every preview update, so redrawing can slow down with many rectangles

### Article

https://note.com/dtp_tranist/n/n789072361c12

---

### Update History

- v1.1.1 (2026-10-11) For horizontal rules, the dialog now starts with the stroke (butt cap) type and keeps the rule's stroke weight. Fixed stroked arrows inheriting the rule's arrowheads and width profile, overlapping anchor points at extreme heights, and missed white fills. The default stroke width and step now suit the stroke unit. Switching between the two stroke types keeps the stroke width
- v1.1.0 (2026-10-10) Added Height. Horizontal rules are now supported. The stroke types now take the thickness as a stroke weight in the Stroke units from Preferences. Added double-headed and reverse icons. A white fill now also falls back to the stroke color
- v1.0.0 (2026-10-10) Initial release

### Script info

- Version: v1.1.1
- First release: 2026-10-10
- Last updated: 2026-10-11
