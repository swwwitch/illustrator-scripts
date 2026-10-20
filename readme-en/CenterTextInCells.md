# Center text in cells divided by rules

[![Direct](https://img.shields.io/badge/Direct%20Link-CenterTextInCells.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/CenterTextInCells.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterTextInCells.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

In a table drawn with an outer frame and rules, centers each text in its cell.
It can also redistribute the rules, evenly or to fit the texts, and set the paragraph alignment to center.

### Features

- Moves each text to the center of the cell formed by the frame and rules (centered by its glyph outlines, both horizontally and vertically)
- "Redistribute vertical rules" / "Redistribute horizontal rules" rearrange the rules inside the frame (a direction with no rules is dimmed)
- "Spacing" chooses how the rules are placed
  - Even: spaces the rules evenly
  - Fit to text: sizes each column (row) to its texts, with equal space on both sides of them
- "Match rule lengths" snaps both ends of redistributed two-point rules to the sides of the frame
- "Align center" sets the paragraph alignment to center
- Preview switch. Closing with Cancel or the Esc key restores both the rules and the texts
- Option (Alt)-click either rule checkbox to toggle both at once
- Texts and rules inside groups and compound paths are included

### Usage

1. Select the texts together with the outer frame and the rules.
2. Run the script. The dialog opens and the preview is shown.
3. Choose how to place the rules and the text options, then confirm with OK.

### Notes

- Cell borders are the vertical and horizontal segments of the selected paths (the four sides of a rectangle and a single rule are treated alike). The nearest border on each side of a text's center forms its cell.
- Texts with no enclosing cell are left in place, and their count is reported at the end.
- Glyph bounds are measured by outlining a duplicate of the text, so ascender and descender space does not affect the result.
- When redistributing rules, the closed path with the largest area in the selection is used as the outer frame. A rule is a thin open path at least 20% as long as the matching side of the frame, with its center inside the frame.
- "Fit to text" uses the widest text in each column. A column with no text gets only the shared space. If the texts do not fit in the frame, the width is divided by the ratio of the text widths with no extra space.
- "Match rule lengths" only affects rules with two anchor points.
- The initial values can be changed with `AVERAGE_RULES`, `DISTRIBUTE_MODE`, `MATCH_LINES_TO_RECT`, `CENTER_IN_CELLS` and `JUSTIFY_CENTER` at the top of the script.

### Update History

- v1.0.0 (2026-10-20) Initial release
