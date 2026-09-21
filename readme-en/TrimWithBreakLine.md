# Trim an image with break lines

[![Direct](https://img.shields.io/badge/Direct%20Link-TrimWithBreakLine.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/TrimWithBreakLine.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TrimWithBreakLine.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Select an object (image, group, path, …) and a path drawn over it: the script builds the two parts without the covered area and closes them up to a set gap.
- A path spanning the artwork horizontally splits it top and bottom, one spanning it vertically splits it left and right (decided from how much of the artwork it spans).
- When the path reaches an edge of the artwork (its top edge, say), the script keeps that side only and trims the rest away, leaving a single part.
- Handy for shortening long screenshots or menus while keeping both ends visible.
- The cut edge can be bent with a Flag, Rise or Rise (straight) warp, and break lines can be drawn along it.

<img alt="The Trim and Break Line dialog" src="../png/ss-882-728-144-20260921-051717.png" width="50%" />

### Features

- Trims the artwork to the width of the mask path (its height when cutting left and right)
- Picks the cut direction automatically (top/bottom or left/right)
- Cut-edge shape from Flag, Rise or Rise (straight), with the bend set in percent
- Both cut edges come from the same curve, so the gap stays even after closing up
- Keeps one side only, with a single break line, when the path covers an edge of the artwork
- Gap between the parts, entered in the ruler unit
- Break lines along the cut edge (solid or dashed, segment count, cap)
- Groups each break line with the part it belongs to
- Live preview from the moment the dialog opens
- Restores the settings used last time (`~/Library/Application Support/TrimWithBreakLine/settings.txt`)
- Distances use the ruler unit, the stroke weight uses the Stroke unit (Preferences › Units)

### How to use

1. Draw a rectangle (or any path) over the area to drop, then select it together with the artwork (image, group, path, …). To keep one side only, draw it over the area to keep, reaching past that edge of the artwork.
2. Run the script.
3. Set the cut-edge shape, the gap and the break lines in the dialog while watching the preview.
4. Click OK. The original artwork and the mask path are removed, and the new parts are selected.

The path acts as the mask shape, and its size across the cut sets the finished size (when the artwork is a path too, the one in front is the mask). The cut direction follows how much of the artwork's width and height the path spans.

Drawn inside the two edges along the cut (the top and bottom edges when cutting top and bottom), it leaves two parts with that area dropped. Drawn over one of those edges, it leaves the covered side only as a single part, and the Gap field goes unused. Covering both edges leaves nothing to keep, so the script stops.

### Dialog

| Item | Description |
| --- | --- |
| Height / Width | Size of the mask shape along the cut (%). 100% keeps it as drawn; the label follows the direction |
| Offset | Position of the mask shape (ruler unit). Positive moves it down or right, negative up or left |
| Warp style | Shape of the cut edge: Flag waves, Rise curves upward to the right, Rise (straight) slants in a straight line |
| Bend | How much the cut edge bends (-100 to 100%). 0 cuts along a straight line; negative values bend it the other way |
| Gap | Distance between the two parts after closing up (ruler unit). Unused when only one side is kept |
| Add break lines | Draws a break line along each cut edge |
| Line style | Solid or dashed |
| Segments | Number of dashes. Dash and gap are equal, and both ends finish with a dash |
| Weight | Stroke weight of the break lines (Stroke unit) |
| Cap | None (butt cap) or Round |
| Group with parts | Groups each break line with the part it was cut from |

### Keyboard

| Key | Action |
| --- | --- |
| Up / Down | Change a number field by 1 |
| Shift + Up / Down | Snap to multiples of ten |
| Option + Up / Down | Change by 0.1 |

### Notes

- Select exactly two objects: the artwork and one path.
- The mask path must leave at least one of the two edges along the cut direction clear.
- The tone of the lines is set by `RULE_STROKE_GRAY` at the top of the script.
- Clicking OK stores the settings and the next run starts from them.
- The preview is built from duplicates; Cancel removes them and restores the original selection.

### Article

- [DTP Transit (Japanese)](https://note.com/dtp_tranist/n/n2483bd96e284)

### Update history

- v1.0.6 (20260921) : When cutting left and right, the left part now stays in place and the right part closes up to it
- v1.0.5 (20260921) : Bend now accepts negative values (-100 to 100%) to bend the cut edge the other way
- v1.0.4 (20260921) : A path covering an edge of the artwork now keeps that side only. Added the Rise (straight) cut-edge style. Stacked the radio buttons (style, line style, cap) and shortened the number fields by one character
- v1.0.3 (20260921) : The cut direction is now decided from how much of the artwork the path spans. Unified the wording around "break lines" and matched the tooltips for weight, gap and offset to the actual unit and direction
- v1.0.2 (20260921) : Left/right cutting with automatic direction detection; any object can be the artwork; settings are restored. Added the Mask shape panel (height and offset) and the stroke weight of the break lines; distances use the ruler unit and the weight uses the Stroke unit
- v1.0.1 (20260921) : Initial release
