# Re-lay out artboards so the whole grid is near-square

[![Direct](https://img.shields.io/badge/Direct%20Link-ArrangeArtboardsToSquareGrid.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ArrangeArtboardsToSquareGrid.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeArtboardsToSquareGrid.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Re-lays out every artboard so the whole grid's outline is as close to a square as possible.

### Features

- Each artboard's artwork moves with it
- The grid is centered on the canvas

### Usage

1. Open the document.
2. Run the script.

### Notes

- Use GridArrangeArtboards.jsx to derive the grid from the artboard names instead.

### Update History

- v1.0
- v1.1.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.1.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.2 (2026-09-28): In Japanese, the artboard count and recommended columns now use a full-width colon. The button row is now built with the shared part
- v1.1.3 (2026-09-29): Dialog opacity changed to 98%
- v1.1.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.1.5 (2026-09-30): Button rows with only right-side buttons are now centered
- v1.1.6 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.1.7 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.1.8 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
