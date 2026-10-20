# Tidy up a ruled table and center texts in its cells

[![Direct](https://img.shields.io/badge/Direct%20Link-ArrangeLikeTableCells.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/ArrangeLikeTableCells.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeLikeTableCells.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

In a table drawn with an outer frame and rules, adjusts the frame size, the rule spacing and the stroke weights, then centers each text in its cell based on its glyph outlines.
It can also round the frame corners and fill the cells with random gray tints.

### Features

#### Frame

- "Width" / "Height" change the size of the frame (prefilled with its current size; the frame and rules are scaled from the top-left corner, keeping stroke widths and text sizes)
- "Convert frame to rectangle" replaces the four rules forming the frame with a single rectangle in the same place (the stroke is taken from the top rule)
- "Round corners" applies the Round Corners effect with the given radius to the frame rectangle

#### Rules

- "Adjust vertical rule spacing" / "Adjust horizontal rule spacing" adjust the positions of the rules inside the frame (a direction with no rules is dimmed)
- "Spacing" chooses how the rules are placed
  - Even: spaces the rules evenly
  - Fit to text: sizes each column (row) to its texts, with equal space on both sides of them
- "Match rule lengths to frame" snaps both ends of adjusted two-point rules to the sides of the frame
- "Match stroke weights" sets the stroke weight of the frame and rules to the given value (prefilled with the most common weight)

#### Text

- "Center in cells" moves each text to the center of the cell formed by the frame and rules (centered by its glyph outlines, both horizontally and vertically)
- "Align center" sets the paragraph alignment to center

#### Cells

- "Create cell shapes" adds a rectangle filled with gray in each cell, sent to the back
- "Fill" chooses how the cell shapes are filled
  - Same tint (K15): every cell is filled with K15
  - Random tints: each cell gets a random tint from K10 to K60
- "Group" puts the cell shapes in one group (when off, they stay separate)

#### General

- Numeric fields step with the up/down buttons and arrow keys (Shift snaps to 10s, Option steps by 0.1). Width, height and radius use the ruler units; the stroke weight uses the stroke units
- The preview updates while you type
- The preview is always on. Closing with Cancel or the Esc key restores everything
- Option (Alt)-click either rule checkbox to toggle both at once
- Texts and rules inside groups and compound paths are included

### Usage

1. Select the texts together with the outer frame and the rules (the frame and rules alone also work).
2. Run the script. The dialog opens and the preview is shown.
3. Set the frame, rule, text and cell options, then confirm with OK.

### Notes

- Cell borders are the vertical and horizontal segments of the selected paths (the four sides of a rectangle and a single rule are treated alike). The nearest border on each side of a text's center forms its cell.
- Texts with no enclosing cell are left in place, and their count is reported at the end.
- Glyph bounds are measured by outlining a duplicate of the text, so ascender and descender space does not affect the result.
- The outer frame is the closed path with the largest area in the selection. Even when the rules extend beyond the rectangle, the rectangle defines the cell area. Without a closed path, a frame drawn with four rules is used (left and right at the outermost vertical rules, top and bottom at the outermost horizontal rules); those four rules are not moved when adjusting the spacing, get projecting caps so the corners close, and "Match rule lengths to frame" also trims them to run from corner to corner.
- A rule is a thin open path at least 20% as long as the matching side of the frame, lying entirely inside the frame across its stroke (a line on a frame edge counts as the edge). Lines tilted by up to about 1 degree are treated as vertical or horizontal.
- Changing "Width" / "Height" scales the anchor points of the frame and rules from the top-left corner. Texts are not scaled; with "Center in cells" on, they move to the centers of the resized cells.
- "Fit to text" uses the widest text in each column. A column with no text gets only the shared space. If the texts do not fit in the frame, the width is divided by the ratio of the text widths with no extra space.
- "Match rule lengths to frame" only affects rules with two anchor points.
- "Convert frame to rectangle" is available only when there is no closed path and the frame is drawn with four rules. The rectangle has no fill and takes the stroke color, weight and dashes from the top rule (if the four sides differ in color, the top one is used).
- "Round corners" is available when the frame is a rectangle (a closed-path frame, or with "Convert frame to rectangle" on). It is applied as an effect, leaving the path shape unchanged. Cell fills are rounded by the same radius at the frame corners only.
- "Match stroke weights" only changes paths that have a stroke.
- Cell shapes use a K-only CMYK color in CMYK documents and grayscale in RGB documents. With "Random tints", each cell gets a random tint that stays the same while the preview updates. The fills are sent to the back of the layer (or group) holding the backmost selected item. If the frame rectangle has a fill, the cell fills are hidden behind it.
- The initial values can be changed with `ADJUST_RULE_SPACING`, `DISTRIBUTE_MODE`, `MATCH_RULE_LENGTHS`, `CONVERT_FRAME_TO_RECT`, `ROUND_CORNERS`, `CORNER_RADIUS_PT`, `MATCH_STROKE_WIDTHS`, `CREATE_CELL_SHAPES`, `CELL_FILL_MODE`, `FILL_UNIFORM_TINT`, `GROUP_CELL_FILLS`, `FILL_TINT_MIN`, `FILL_TINT_MAX`, `CENTER_IN_CELLS` and `JUSTIFY_CENTER` at the top of the script.

### Article

[Introduction article (Japanese)｜DTP Transit 別館](https://note.com/dtp_tranist/n/n22aa0e767930)

### Update History

- v1.3.1 (2026-10-20) The script now runs with only the frame and rules selected, without texts (the Text panel and "Fit to text" are disabled)
- v1.3.0 (2026-10-20) The Cells panel now has "Create cell shapes" with a choice of fill ("Fill": "Same tint (K15)" / "Random tints" (K10-K60)). A frame drawn with four rules now gets projecting caps automatically. The dialog now has two columns (Rules and Text on the left, Frame and Cells on the right). Removed the Preview checkbox; the preview is now always on
- v1.2.0 (2026-10-20) Added the Frame panel with "Width" / "Height" (scale the frame and rules from the top-left corner), "Convert frame to rectangle" (replaces the four rules forming the frame with a single rectangle in the same place) and "Round corners" (applies the Round Corners effect to the frame rectangle). Added "Match stroke weights" to the Rules panel, and "Fill with random tints" (a random gray from K10 to K60 in each cell) and "Group" to the Cells panel. Numeric fields now have up/down buttons, and the preview updates while you type. Renamed the options to "Adjust vertical rule spacing", "Adjust horizontal rule spacing" and "Match rule lengths to frame", and the dialog title to "Arrange Like a Table". The dialog now opens where it was last closed instead of moving sideways to avoid the selection. Fixed the rectangle not being used as the frame when the rules extend beyond it (a closed path, when present, is now always the frame). Slightly tilted lines (up to about 1 degree) now count as cell borders and as rules forming a frame. Lines lying on a frame edge are no longer redistributed as rules
- v1.1.0 (2026-10-20) Rules can now be redistributed in a frame drawn with four rules instead of a rectangle (the positions of the outermost rules form the frame, and "Match rule lengths" also trims the four frame rules to run from corner to corner)
- v1.0.0 (2026-10-20) Initial release
