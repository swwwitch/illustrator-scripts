# SetStrokeAndArrowheads

[![Direct](https://img.shields.io/badge/Direct%20Link-SetStrokeAndArrowheads.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/SetStrokeAndArrowheads.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SetStrokeAndArrowheads.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Sets stroke width and arrowheads (shape, scale and tip alignment for both ends) of the selection at once.
- Arrowheads are not exposed to the Illustrator DOM, so a temporary action (ai_plugin_setStroke) is generated and played.
- The dialog supports preview: stroke width updates through the DOM instantly, arrowhead settings are previewed by playing the action and undoing it.

### Process Flow

1. Verify that a document is open and objects are selected
2. Enter stroke width, arrowheads and tip alignment in the dialog (with preview)
3. Build an .aia (action) source from the entered values
4. Write it as a temporary file, load, play, then discard it

### Notes

- Arrowhead and tip alignment names must match the Illustrator UI labels (language dependent).
- The arrowhead scale keys (asc1 / asc2) are estimated values.

### Script info

- Version: v1.0.0
- First release: 2026-07-22
- Last updated: 2026-07-22
