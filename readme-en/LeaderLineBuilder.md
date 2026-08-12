# LeaderLineBuilder

[![Direct](https://img.shields.io/badge/Direct%20Link-LeaderLineBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/LeaderLineBuilder.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LeaderLineBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Builds leader lines at a specified angle from the bounding box of the selected paths or groups.
- Each leader line is a three-point shape: a diagonal segment plus a horizontal segment.
- Angle, direction, line style, tip marker and edge are adjusted with a live preview in the dialog.
- The selected object is replaced by the leader line (the shape acts as a construction guide).

### How to use

1. Select a path (2 or more points) or a group to use as the reference.
2. Run the script and adjust the settings in the dialog.
3. Click OK. The selected object is replaced by the leader line.

Selecting a text frame together with the object switches to text alignment mode. The direction is detected from the relative position of the object and the text, the horizontal end is aligned to the matching edge of the text, and the bend point is offset from the text by the given distance.

### Dialog settings

| Item | Description |
| --- | --- |
| Apply Scope | "Update all" applies every setting. "Keep each direction" keeps the direction stored in each selected leader line |
| Angle | Angle of the diagonal segment (greater than 0, less than 90). Presets: 30° / 45° / 60° |
| Diagonal Direction | Where the leader line runs (upper left, lower left, upper right, lower right) |
| Line Style | Line color (black / white / custom), stroke width, line end shape (none / round) |
| Tip Marker | A circle or an arrow at the tip. Filled or outlined, with a size |
| Group items | Groups the line, tip marker and edge together |
| Edge | Places a thicker copy of the line and tip marker behind them. The color can be chosen |
| Text | Distance to the text (pt). Available only when a text frame is also selected |
| Zoom | Changes the view zoom while the dialog is open. "Defer redraw" skips redrawing while dragging |

Stroke width and tip marker size follow the stroke unit in your preferences. The distance to text is always in points.

In numeric fields the arrow keys step the value. Shift steps larger, Option (Alt) steps finer.

### Re-applying

Leader lines created with "Group items" store the bounding box and direction of the original object as tags. Selecting such a leader line and running the script again rebuilds it at the same position from the stored information.

When several leader lines are selected at once, "Keep each direction" is preselected so that everything except the direction (color, stroke width, tip marker and so on) can be updated in one pass.

### Process flow

1. Verify that a document is open and objects are selected
2. Collect the targets (paths, groups) and text frames from the selection
3. Measure the bounding box and stroke attributes of each target
4. Compute the three points from the dialog settings and draw the edge, line and tip marker (preview)
5. On OK, remove the original object and replace it with the leader line

### Notes

- The original object is deleted. Duplicate it first if you need to keep it.
- The angle must be greater than 0 and less than 90; other values cannot be confirmed.
- With "Group items" off, no tags are written, so the result cannot be re-applied later.
- An arrow tip marker is always filled; the outlined option is not available for it.
- Colors are created to match the document color space (CMYK or RGB).

### Changelog

- v1.5.2 (2026-08-12) : Fixed the tip marker being placed on the wrong end in text alignment mode, fixed "Group items" being unavailable when the edge is on without a tip marker, reorganized the internal structure

### Script info

- Version: v1.5.2
- First release: 2026-03-06
- Last updated: 2026-08-12
- Article: https://note.com/dtp_tranist/n/n506df641d5c5
