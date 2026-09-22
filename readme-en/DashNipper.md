# Split a dashed line into separate segments

[![Direct](https://img.shields.io/badge/Direct%20Link-DashNipper.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/DashNipper.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashNipper.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- With a dashed line selected, the script outlines it, splits it into its dashes, replaces each dash with a line that looks the same (its center line), and groups the lines of each dashed line. The result looks like the original dashed line, but every dash is a separate path.
- If the selection contains no dashed line, the script shows an alert and stops.

### Main Features

- Caps: dashes from a butt-cap dashed line get butt caps; dashes from a round-cap dashed line (pill shapes with semicircular ends) get round caps.
- Stroke width and color: the stroke width is taken from the original dashed line, and the stroke color from the fill of the outlined dash.
- Curves: dashes on a curved path become center lines that follow the curve. A dash that crosses an anchor point or a mitered corner becomes a center line with an anchor point in the middle.
- Direction: the direction is chosen so that the thickness matches the original stroke width, so dashes shorter than the stroke width are not turned sideways.
- Grouping: the lines made from each dashed line are grouped together. Dashes that could not become center lines go into the same group as outlines. A single line is not grouped.
- Stacking order: each group is placed where the dashed line was (same layer or group), and the new groups are selected afterward.
- Multiple dashed lines: the selected dashed lines (including those inside groups) are processed one at a time. The stroke width is read from each one.
- Retry: if some dashes could not become center lines, only those dashes are left selected. Run the script again as is to convert them with the original stroke width.

### Usage

1. Select one or more dashed lines (selecting a group that contains dashed lines also works)
2. Run the script
3. If the alert about dashes that could not be converted appears, run the script again with those dashes still selected

### Notes

- Only paths with Dashed Line set in the Stroke panel are processed. Dashes made with a brush or with a second stroke in the Appearance panel, and dashed compound paths, are not.
- Hidden or locked paths, clipping paths, and guides inside groups are skipped.
- If a dashed line has a fill, the fill is deleted (otherwise it would stay in front of the lines and hide them).
- Other selected objects, including non-dashed objects inside selected groups, are left as they are.
- The following dashes cannot become center lines and stay outlined (an alert reports how many):
  - Zero-length dashes with round caps (the dots of a dotted line)
  - Dashes that cross a beveled or rounded corner
  - Dashes whose two sides have different numbers of anchor points
  - Dashes whose thickness changes with a width profile
- Dashes that could not become center lines keep the original stroke width in their note in the Attributes panel (`DashNipper:strokeWidth=…`). A retry reads this note, so do not delete it.
- On a closed dashed path, outlining can leave very short segments and tiny fragments near the seam. The short segments are cleaned up before the center line is computed, and the fragments are deleted.
- Dashes with projecting caps become lines with butt caps (the visible length stays the same).

### Article

- [DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/nae6882ac8a73)

### Update History

- v1.0.0 (20260922) : Initial release
- v1.1.0 (20260922) : The lines of each dashed line are now grouped. Dashed lines inside groups are now processed; hidden or locked paths, clipping paths, and guides are skipped. The script now shows an alert and stops when no dashed line is selected, and converting a single already-outlined shape is no longer supported. Fixed dashes near the seam of a closed dashed path failing to convert; tiny fragments left by outlining are now deleted. Fixed the fill of a filled dashed line staying in front of the lines (the fill is now deleted). Fixed dashes shorter than 5% of the stroke width being deleted. Fixed the direction and stroke width sometimes being off when the dash length is close to the stroke width. Fixed some dashes of a thin round-cap dashed line not becoming center lines. Only the dashes that could not be converted are now left selected, and running the script again converts them with the original stroke width
- v1.1.1 (20260922) : Fixed dashes near the seam of a round-cap dashed line, such as a circle, sometimes not becoming center lines
