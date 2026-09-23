# Place objects at the center of each grid cell

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartObjectDistributor.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartObjectDistributor.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectDistributor.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

![ss-2006-988-72-20250520-163549](https://github.com/user-attachments/assets/b438fa19-3d35-4b65-9ebc-c5de37946511)

## Overview

This script places the selected objects at the center of each cell of a grid with the given number of rows and columns.

The placement area can be the current artboard, the backmost object, or a rectangle in the "_target" layer. The cell rectangles can be kept as they are, converted to guides, or turned into artboards. Divisions, margins, and gutters are adjusted with a live preview.

## Main features

- Placement area switching (current artboard / backmost object / rectangle in the "_target" layer)
- Rows, columns, gutter, and margin settings
- Initial rows and columns derived from the selection count and the area's aspect ratio, so cells stay nearly square
- Cell handling: Keep as Rectangle, Convert to Guides, or Convert to Artboards
- Cell fill (black, white, or no fill) and opacity
- Shuffle button to randomize the order in which objects fill the cells
- Objects beyond the number of cells are parked clear of other objects and artboards
- Transparency grid toggle, restored to its original state on exit
- Arrow-key stepping (Shift: by 10, Option: by 0.1)
- Live preview updated as values change
- Values are entered in Illustrator's ruler unit
- Japanese and English UI

## Usage

1. Select the objects to arrange in a grid.
2. Run `SmartObjectDistributor.jsx`.
3. Configure the settings in the dialog panels.
   - Placement Area: the region the grid is laid out in
   - Divisions & Margin: rows, columns, gutter, margin
   - Cell Handling: what happens to the rectangles, color, opacity
4. Click Shuffle to change the assignment order if needed.
5. Adjust the values while checking the preview.
6. Click OK to apply.

## Panels and settings

| Panel | Settings |
| --- | --- |
| Placement Area | Current artboard / backmost object / rectangle in the "_target" layer |
| Divisions & Margin | Rows, columns, gutter, margin |
| Cell Handling | Keep as Rectangle / Convert to Guides / Convert to Artboards, color (black / white / no fill), opacity, transparency grid |

![ss-746-988-72-20250521-053754](https://github.com/user-attachments/assets/3bec4d25-184d-46ab-98c5-fecbb2be9c0d)

## Notes

- The rectangle in the "_target" layer is hidden while the script runs and is never placed into a cell. A rectangle you had already hidden stays hidden afterwards.
- When the "_target" layer holds several paths, a rectangle wins; open lines and paths with no area are never used as the placement area.
- When Backmost Object is selected, that framing object is excluded from the objects to place.
- Objects beyond the number of cells are parked where they overlap neither other objects nor artboards.
- Any existing "cell-background" layer is hidden while the dialog is open and removed when you click OK, which clears the result of the previous run. Cancel restores its original visibility.
- With Convert to Guides or Convert to Artboards, the cell color is fixed to No Fill.
- Up to 100 rows and 100 columns, and up to 1000 cells. Out-of-range values are clamped when the field commits.
- Nothing is drawn when the margin or gutter leaves no room for cells, and clicking OK keeps the dialog open and explains why.
- Each preview redraw undoes the previous one, so changing settings does not keep growing Illustrator's edit history (your pre-run history is not pushed out). Anything Undo fails to revert is cleaned up by restoring the recorded centers and dropping the preview layer.
- Besides Cancel, pressing Esc or closing the window also discards the preview and restores the original state.
- With Keep as Rectangle, the cell rectangles are selected after clicking OK.
- Values follow Illustrator's ruler unit setting.

## Article

[Arrange objects in a grid with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/na3c45cea09b7)

## Changelog

- v1.9.7 (2026-09-23): Fixed Cancel leaving the "cell-background" layer deleted (it is now removed only on OK). A "_target" rectangle that was already hidden is no longer revealed. Fixed OK closing silently when the settings cannot form a grid; the dialog now stays open and explains why. Field values are clamped on commit (no negative margins, no opacity above 100%). Fixed re-clicking the selected color radio resetting a typed opacity. The "_target" layer now prefers a rectangle and never uses open lines or paths with no area. Unit code 5 is labelled H for distances and Q for type sizes. The live preview is throttled while typing on large grids. The default division no longer exceeds the cell cap
- v1.9.6 (2026-09-19): Unified unit handling into one table covering every ruler unit (m, yd, ft, pica, and so on). Split out the layout constants and tidied naming and JSDoc
- v1.9.5 (2026-07-27): Preview rollback now uses app.undo(), so the edit history no longer grows on every change; whatever Undo cannot revert is cleaned up by restoring the centers and dropping the preview layer. Added cleanup on Esc and window close. Fixed Undo reverting the "_target" rectangle's visibility toggle instead of the preview when switching the placement area
- v1.9.0 (2026-07-27): Made the preview independent of Undo (fixes objects changing when a button is clicked repeatedly), added input validation and limits, reorganized UI wording and tooltips
- v1.8.0 (2026-04-30): Cell handling as radio buttons, placement separated from cell drawing, transparency grid state restored
- v1.0.0 (2025-06-05): Initial version
- v0.5.8 (2025-06-05): Added temporary artboard usage from the "_target" layer
- v0.5.7 (2025-06-05): Immediate preview update for margin changes
- v0.5.6 (2025-06-05): Improved UI control
- v0.5.5 (2025-06-05): Simplified grid settings, added previous cell removal
- v0.5.4 (2025-06-05): Refined UI and label structure, auto-remove cell-background layer
- v0.5.3 (2025-06-05): Unified gutter and margin settings
