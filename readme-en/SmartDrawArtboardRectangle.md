# SmartDrawArtboardRectangle

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartDrawArtboardRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/SmartDrawArtboardRectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDrawArtboardRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Draws rectangles that match the active or all artboards, with optional offset.
- Supports color (None / K100 15% / HEX / CMYK), stacking order (Front / Back / bg layer), and live preview.
- Optional post-draw actions: "Make guides" and "Convert to Live Shape".
- Always shows the center widget (center point) on live shapes.

### Key Features

- Offset with Bleed presets (3mm / 12H / 0.125in)
- Color modes (None / K100 15% / HEX / CMYK)
- Z-order (Front / Back / bg layer; defaults to Bring to Front)
- Target scope (Current artboard / All artboards)
- Options (Make guides: default OFF / Convert to Live Shape: default ON)
- Always shows the center widget (applied via a recorded action)
- Preview (1pt dashed stroke, 50% tone, dedicated layer)
- Hotkeys (F: Front / B: Back / L: bg layer / G: Make guides / C: Current artboard / A: All)
- Dialog initial position & opacity settings

### Update History

- v1.0 (20250820): Initial version
- v1.5.1 (20250824): Added post-draw options (Convert to Live Shape, Make guides)
- v1.5.2 (20260531): Always show center widget, unified unit table, CMYK input fix, hotkey rework
- v1.5.3 (20260531): Renamed object to "<Rectangle>", tweaked offset field width
- v1.5.4 (20260601): Fixed Error 8705 when the front-most layer is a template/locked layer
- v1.5.5 (20260625): Reliable preview-layer cleanup (no leftover/mis-draw), guide mode renames to "<Guide>" and clears selection, HEX field accepts color names / #RGB shorthand / grayNN, sturdier hotkey guard, center widget only on live shapes, internal refactoring; latest version

### Script info

- Version: v1.5.5
