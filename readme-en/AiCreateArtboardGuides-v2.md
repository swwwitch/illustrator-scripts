# AiCreateArtboardGuides-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-AiCreateArtboardGuides--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/AiCreateArtboardGuides-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCreateArtboardGuides-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A tool that organizes and creates guides based on artboards. Three groups of settings are configured in a single dialog.

- Convert ruler guides: detects ruler guides that overlap an artboard and redraws them as straight, artboard-based guides (master ON/OFF)
- Center guides: creates vertical and horizontal guides at the center of each artboard
- Edge guides: creates guides along the top, bottom, left and right of each artboard (master ON/OFF, off by default)
- Center and edge guides can be created even when the document has no guides at all
- Every created guide is collected on a `_guide` layer (created automatically if missing; unlocked and shown while in use)

Settings

- "Extend outward" sets how far the guides run past the artboard edge
- "All artboards" is set separately for conversion (every overlapping artboard) and for center/edge guides (all artboards or the active one only)
- Input values are read in the current ruler unit and converted to points internally (the unit comes from the `rulerType` preference)

Preview

- Live preview follows every change: coloured provisional lines are drawn on a dedicated layer and replaced by real guides on OK

### Changing values

- Up/Down arrow keys step by ±1
- Hold Shift to step by ±10

### Article

https://note.com/dtp_tranist/n/n56d9c936a364

### Script info

- Version: v1.1.0
