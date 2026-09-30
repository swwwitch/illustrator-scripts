# Swap the appearance or the text between two objects

[![Direct](https://img.shields.io/badge/Direct%20Link-SwapStyle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/SwapStyle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapStyle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Swaps the appearance, or the text content, between two selected objects.

### Features

- The dialog switches between swapping style and swapping text
- Style swapping combines three groups of properties
  - Graphic style (the whole current appearance)
  - Basic fill and stroke (fill color, stroke color, stroke weight)
  - Character attributes

### Usage

1. Select the two objects you want to swap.
2. Run the script.
3. Choose what to swap and click OK.

### Notes

- Nothing happens unless exactly two objects are selected.

### Update History

- v1.1.0 (2026-05-23)
- v1.1.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.2 (2026-09-28): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure. The button row is now built with the shared part
- v1.1.3 (2026-09-29): Dialog opacity changed to 98%
- v1.1.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.1.5 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.1.6 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
