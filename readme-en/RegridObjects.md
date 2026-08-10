# RegridObjects

[![Direct](https://img.shields.io/badge/Direct%20Link-RegridObjects.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RegridObjects.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegridObjects.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Assumes the selected objects are roughly arranged in a grid, and re-lays them out using horizontal and vertical spacing values.
- Always-on preview. Values are typed directly into the fields, or stepped with Up/Down (×10 with Shift, ×0.1 with Option).
- Spacing is entered in the current ruler unit (mm / pt / px, and so on) and converted to points internally; the unit is shown in the panel title.
- Existing groups (including clip groups) are treated as a single object with one bounding box, rather than being broken apart.
- Objects are not grouped automatically afterwards; they simply stay selected.
- The dialog switches between Japanese and English automatically (`$.locale`).
- Link mirrors the horizontal value into the vertical one.
- Brick: offsets every other row horizontally by half a pitch.
- Honeycomb: used together with Brick, it shifts odd rows by half of (width + horizontal spacing) and scales the row height to 0.75, producing a honeycomb layout (the vertical value still applies).
- Force grid: instead of inferring columns and rows from proximity, it assigns (row, column) top to bottom and left to right.
- Center (a sub-option of Force grid): centers each object within its cell (column width × row height).
- Transpose is a toggle: on, it swaps rows and columns while tolerating gaps; off, it returns to the pre-transpose state.
- Transposing a single row into a single column, and vice versa, is supported.

### Update History

- v1.6.0 (2026-07-08): Added Center (a sub-option of Force grid), made Transpose a toggle that reverts when off, added ruler-unit input (mm / pt / px, converted to points internally), and tidied the apply functions and their naming
- v1.0 (2025-10-31): Always-on preview, linked values (vertical dimmed), and arrow-key stepping

### Article

https://note.com/dtp_tranist/n/n08861d0e40c3

### Script info

- Version: v1.6.0
