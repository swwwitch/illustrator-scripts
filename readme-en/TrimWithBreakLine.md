# Trim an image with break lines

[![Direct](https://img.shields.io/badge/Direct%20Link-TrimWithBreakLine.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/TrimWithBreakLine.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TrimWithBreakLine.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Select an image and a band-shaped path drawn over it: the script builds the upper and lower parts without the banded area and closes them up to a set gap.
- Handy for shortening long screenshots or menus while keeping both ends visible.
- The cut edge can be bent with a Flag or Rise warp, and break lines can be drawn along it.

### Features

- Trims the image to the width of the band path (the sides are cropped too)
- Cut-edge shape from Flag or Rise, with the bend set in percent
- Both cut edges come from the same curve, so the gap stays even after closing up
- Gap between the parts, entered in the ruler unit
- Break lines along the cut edge (solid or dashed, segment count, cap)
- Groups each break line with the part it belongs to
- Live preview from the moment the dialog opens

### How to use

1. Draw a rectangle (or any path) over the area to drop, then select it together with the image (placed or embedded).
2. Run the script.
3. Set the cut-edge shape, the gap and the break lines in the dialog while watching the preview.
4. Click OK. The original image and the band path are removed, and the new parts are selected.

The path in front acts as the band, and its left and right edges set the finished width. Draw it inside the top and bottom edges of the image.

### Dialog

| Item | Description |
| --- | --- |
| Style | Shape of the cut edge: Flag waves, Rise curves upward to the right |
| Bend | How much the cut edge bends (0-100%). 0 cuts along a straight line |
| Gap | Distance between the parts after closing up (ruler unit) |
| Add rules | Draws a break line along each cut edge |
| Style | Solid or dashed |
| Segments | Number of dashes. Dash and gap are equal, and both ends finish with a dash |
| Cap | None (butt cap) or Round |
| Group with parts | Groups each break line with the part it was cut from |

### Keyboard

| Key | Action |
| --- | --- |
| Up / Down | Change a number field by 1 |
| Shift + Up / Down | Snap to multiples of ten |
| Option + Up / Down | Change by 0.1 |

### Notes

- Select exactly two objects: one image and one path.
- The band must sit inside the top and bottom edges of the image.
- Stroke weight and tone are set by `RULE_STROKE_WIDTH` and `RULE_STROKE_GRAY` at the top of the script.
- The preview is built from duplicates; Cancel removes them and restores the original selection.

### Article

- [DTP Transit (Japanese)](https://note.com/dtp_tranist/n/n2483bd96e284)

### Update history

- v1.0.1 (20260921) : Initial release
