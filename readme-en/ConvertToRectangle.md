# ConvertToRectangle

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/ConvertToRectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Updated: 2026-05-24
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

### Script info

- Version: v1.1.0
