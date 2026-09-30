# Create rectangles matching the bounds of the selection

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/ConvertToRectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Creates rectangles matching the bounds of the selected objects
- The unit of creation is either per object or the whole selection
- Margins (in ruler units, negative values inset), a corner-radius live effect, and fill and stroke presets can be set
- The original objects can be kept, turned into a clipping mask, or deleted
- Preview supported: while the dialog is open the selection is dimmed to 50% so the result is easy to compare

### Main Features

- Measure by preview bounds, or by outlining the text (available only when text is selected), with margins in ruler units (negative values inset)
- While the dialog is open the selection's opacity drops to 50% for easier previewing (disabled automatically when a preset changes the opacity)
- Fill and stroke presets (stroke weight follows Illustrator's stroke unit), a corner-radius live effect (radius in ruler units), the stacking order and the treatment of the original object (keep / clipping mask / delete)
- Clipping mask is offered only when a linked or embedded image is selected
- Turning it into a clipping mask preserves the original stacking order

### Update History

- v1.1.1 (2026-09-19) Merged the overlapping `長方形に変換.jsx`; this script is a superset
- v1.2.0 (2026-09-27) Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.1 (2026-09-28) The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.2 (2026-09-28) Fixed the ha (H) unit conversion (1 H was treated as 0.25 pt instead of 0.25 mm). Feet, meters and yards are now converted; units shown as "H" and "pica"
- v1.2.2 (2026-09-28) The button row is now built with the shared part. Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held). Clip groups are now measured by their mask
- v1.2.3 (2026-09-29) Dialog opacity changed to 98%
- v1.2.4 (2026-09-30) Fixed an error when running with characters selected by the Type tool

### Script info

- Version: v1.2.3
