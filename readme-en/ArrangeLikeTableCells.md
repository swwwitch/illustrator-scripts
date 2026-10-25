# Tidy up a ruled table and center texts in its cells

[![Direct](https://img.shields.io/badge/Direct%20Link-ArrangeLikeTableCells.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/ArrangeLikeTableCells.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeLikeTableCells.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

In a table drawn with an outer frame and rules, adjusts the frame size, the rule spacing and the stroke weights, then centers each text in its cell, keeping Latin baselines aligned.
It can also round the frame corners and fill the cells with random gray tints.

### Features

#### Size

- "Width" / "Height" change the size of the frame (prefilled with its current size; the frame and rules are scaled from the top-left corner, keeping stroke widths and text sizes)
- "Preview bounds" makes Width and Height include the stroke weight (initially follows Illustrator's "Use Preview Bounds" preference; changing the stroke weight keeps the size including the stroke)

#### Frame

- "Convert frame to rectangle" replaces the four rules forming the frame with a single rectangle in the same place (the stroke is taken from the top rule)
- "Convert frame to four lines" replaces the frame rectangle with four lines (top, right, bottom, left) in the same place (the stroke is taken from the rectangle; the fill is not kept)
- "Round corners" applies the Round Corners effect with the given radius to the frame rectangle

#### Vertical / Horizontal Rule Spacing

- Each panel chooses how the rules inside the frame are placed in its direction (a direction with no rules is dimmed; Option-click to set the other direction the same way)
  - Even: spaces the rules evenly
  - Fit to text: sizes each column (row) to its texts, with equal space on both sides of them
  - Keep: leaves the rules where they are (they are scaled with the frame when its size changes)
- Both ends of adjusted two-point rules are snapped to the sides of the frame

#### Stroke Color and Weight

- "Adjust stroke weight" makes the stroke weights of the frame and rules the same
  - Max: matches the thickest stroke
  - Min: matches the thinnest stroke
  - Average: matches the average stroke weight (the value is shown next to the option)
  - Specified: uses the value in the field (prefilled with the most common weight)
- "Stroke color" applies the swatch color to the strokes of the frame and rules. Click the swatch to pick a color in Illustrator's Color Picker (initially the color of the first stroke; picking a color turns "Stroke color" on)

#### Text

- "Center in cells" moves each text to the center of the cell formed by the frame and rules (horizontally by its glyph outlines, vertically by the band from the baseline to the cap height, so ascenders and descenders do not shift the baseline)
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
- The outer frame is the closed path with the largest area in the selection. Even when the rules extend beyond the rectangle, the rectangle defines the cell area. Without a closed path, a frame drawn with four rules is used (left and right at the outermost vertical rules, top and bottom at the outermost horizontal rules); those four rules are not moved when adjusting the spacing, get projecting caps so the corners close, and are trimmed to run from corner to corner when the rule spacing is adjusted.
- A rule is a thin open path at least 20% as long as the matching side of the frame, lying entirely inside the frame across its stroke (a line on a frame edge counts as the edge). Lines tilted by up to about 1 degree are treated as vertical or horizontal.
- Changing "Width" / "Height" scales the anchor points of the frame and rules from the top-left corner. Texts are not scaled; with "Center in cells" on, they move to the centers of the resized cells.
- "Fit to text" uses the widest text in each column. A column with no text gets only the shared space. If the texts do not fit in the frame, the width is divided by the ratio of the text widths with no extra space.
- Only rules with two anchor points are snapped to the frame. To keep the rule lengths (for example, when the rules are drawn with gaps from the frame), set `MATCH_RULE_LENGTHS` at the top of the script to `false`.
- "Convert frame to rectangle" is available only when there is no closed path and the frame is drawn with four rules. The rectangle has no fill and takes the stroke color, weight and dashes from the top rule (if the four sides differ in color, the top one is used).
- "Round corners" is available when the frame is a rectangle (a closed-path frame, or with "Convert frame to rectangle" on). It is applied as an effect, leaving the path shape unchanged. Cell fills are rounded by the same radius at the frame corners only.
- "Stroke color" and "Adjust stroke weight" only change paths that have a stroke. Max, min and average are taken from the stroke weights of the frame and rules at launch.
- Cell shapes use a K-only CMYK color in CMYK documents and grayscale in RGB documents. With "Random tints", each cell gets a random tint that stays the same while the preview updates. The fills are sent to the back of the layer (or group) holding the backmost selected item. If the frame rectangle has a fill, the cell fills are hidden behind it.
- The initial values can be changed with `RULE_SPACING_MODE`, `CONVERT_FRAME_TO_RECT`, `ROUND_CORNERS`, `CORNER_RADIUS_PT`, `APPLY_STROKE_COLOR`, `MATCH_STROKE_WIDTHS`, `STROKE_WIDTH_MODE`, `CREATE_CELL_SHAPES`, `CELL_FILL_MODE`, `FILL_UNIFORM_TINT`, `GROUP_CELL_FILLS`, `FILL_TINT_MIN`, `FILL_TINT_MAX`, `CENTER_IN_CELLS` and `JUSTIFY_CENTER` at the top of the script.

### Article

[Introduction article (Japanese)｜DTP Transit 別館](https://note.com/dtp_tranist/n/n22aa0e767930)

### Update History

- v1.5.0 (2026-10-25) Added "Convert frame to four lines" (the reverse of "Convert frame to rectangle"; available when the frame is a four-point rectangle)
- v1.4.3 (2026-10-21) Short rules that stick out of the frame or start partway through a cell are now recognized (a rule whose center lay outside the frame was not treated as a rule, and a rule not reaching the text height did not divide cells)
- v1.4.2 (2026-10-21) "Center in cells" now centers vertically by the band from the baseline to the cap height instead of the glyph outlines (baselines in adjacent cells no longer mismatch because of ascenders or descenders, as with "Lorem" and "Lorem ip")
- v1.4.0 (2026-10-20) Added the Stroke Color and Weight panel, replacing "Match stroke weights" with "Adjust stroke weight". Added "Stroke color" (change the stroke color in the Color Picker). The stroke weight can now be set to "Max", "Min", "Average" (the average is shown next to it) or "Specified". Removed "Match rule lengths to frame" from the dialog; it is now always on (configurable with `MATCH_RULE_LENGTHS`). Moved the Text panel to the bottom of the right column. Moved Width and Height from the Frame panel to a new Size panel and added "Preview bounds" (Width and Height include the stroke weight). Replaced the Rules panel ("Adjust vertical / horizontal rule spacing" and "Spacing") with the Vertical Rule Spacing and Horizontal Rule Spacing panels of radio buttons ("Even", "Fit to text", "Keep"), so each direction can be placed differently. Edges are hidden while the script runs (restored on exit)
- v1.3.1 (2026-10-20) The script now runs with only the frame and rules selected, without texts (the Text panel and "Fit to text" are disabled)
- v1.3.0 (2026-10-20) The Cells panel now has "Create cell shapes" with a choice of fill ("Fill": "Same tint (K15)" / "Random tints" (K10-K60)). A frame drawn with four rules now gets projecting caps automatically. The dialog now has two columns (Rules and Text on the left, Frame and Cells on the right). Removed the Preview checkbox; the preview is now always on
- v1.2.0 (2026-10-20) Added the Frame panel with "Width" / "Height" (scale the frame and rules from the top-left corner), "Convert frame to rectangle" (replaces the four rules forming the frame with a single rectangle in the same place) and "Round corners" (applies the Round Corners effect to the frame rectangle). Added "Match stroke weights" to the Rules panel, and "Fill with random tints" (a random gray from K10 to K60 in each cell) and "Group" to the Cells panel. Numeric fields now have up/down buttons, and the preview updates while you type. Renamed the options to "Adjust vertical rule spacing", "Adjust horizontal rule spacing" and "Match rule lengths to frame", and the dialog title to "Arrange Like a Table". The dialog now opens where it was last closed instead of moving sideways to avoid the selection. Fixed the rectangle not being used as the frame when the rules extend beyond it (a closed path, when present, is now always the frame). Slightly tilted lines (up to about 1 degree) now count as cell borders and as rules forming a frame. Lines lying on a frame edge are no longer redistributed as rules
- v1.1.0 (2026-10-20) Rules can now be redistributed in a frame drawn with four rules instead of a rectangle (the positions of the outermost rules form the frame, and "Match rule lengths" also trims the four frame rules to run from corner to corner)
- v1.0.0 (2026-10-20) Initial release
