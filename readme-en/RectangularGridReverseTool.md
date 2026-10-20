# Rebuild horizontal and vertical lines into a grid

[![Direct](https://img.shields.io/badge/Direct%20Link-RectangularGridReverseTool.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/RectangularGridReverseTool.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangularGridReverseTool.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Analyzes the selected horizontal and vertical lines (including an outer-frame rectangle) and rebuilds them into a rectangular grid.

Line ends are snapped to the intersections and rows and columns are evenly spaced, so even tables with merged cells become a regular lattice.

### Features

- Pre-processing: treat an outer-frame rectangle as four separate lines
- Distribution Mode: do not distribute / evenly (force) / evenly + merged cells
  - With merged cells, lines at the same position form one row or column, and line ends snap to the nearest intersection
- Equalize Targets: choose vertical lines (column widths) and horizontal lines (row heights) separately, optionally keeping the first column or first row as it is
  - Option-click Vertical lines or Horizontal lines to set the other one the same way
- Lines: projecting caps, dashed to solid
- Stroke Width: match the maximum, minimum, average, or a specified weight
  - The Specified field steps with the ∧∨ buttons and arrow keys, and accepts arithmetic and unit values (such as "1 mm")
- Post-processing: convert the outer frame to a rectangle, center point text vertically in cells, group the result
- Preview: check the result without closing the dialog

### Usage

1. Select the table rules (horizontal and vertical lines, and an outer-frame rectangle). Lines inside groups are included.
2. Run the script.
3. Set the options, check the preview, and click OK. Cancel restores the original.

### Notes

- Paths with a fill and lines that do not have exactly two anchor points are ignored.
- At least two horizontal and two vertical lines are required.
- Center point text vertically in cells moves every point text inside the grid, whether or not it is selected.

### Article

[Build a rectangular grid from uneven rules with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n4f290b420561)

### Update History

- v1.3.9 (2026-10-20) Fixed the preview not updating when only lines are selected (Split outer frame is dimmed when there is no rectangle). Fixed Center point text vertically using the rows from before alignment, and a possible stop on locked or hidden text. Splitting a rectangle and cancelling now keep its stacking order and attributes. The Specified stroke width now starts at 0.01
- v1.3.8 (2026-10-20) Fixed the Distribution Mode and Stroke Width choices sometimes being read incorrectly. The stroke width field now accepts arithmetic and unit values and reverts non-numbers. Renamed the "Line (post-process)" panel to "Lines" and noted the Option-click on Vertical/Horizontal lines in the tooltips
- v1.3.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.3.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.3.5 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.3.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.3.3 (20260929): Dialog opacity changed to 98%
- v1.3.2 (20260928): The button row is now built with the shared part
- v1.3.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.0 (20260430): Added Lock first column / Lock first row to Equalize Targets. Paths with a fill are now excluded. Dialog title changed to "Rebuild Grid"
- v1.1.0 (20260429): Added Equalize Targets (vertical / horizontal lines). Lines inside groups are now included, and an outer-frame rectangle can be split into four lines
- v1.0.0 (20260427): Initial release
