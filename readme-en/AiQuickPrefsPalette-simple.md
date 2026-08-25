# AiQuickPrefsPalette-simple

[![Direct](https://img.shields.io/badge/Direct%20Link-AiQuickPrefsPalette--simple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/AiQuickPrefsPalette-simple.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiQuickPrefsPalette-simple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent-palette utility for batch-toggling various Illustrator preferences and flipping/rotating the selection. Every action applies immediately when triggered.

- Runs in a persistent-engine palette; writes and DOM operations are delegated to the main engine via BridgeTalk (reads are fetched directly/synchronously)
- Top is a two-column row (left = Key input / Align Options / Transform Options, right = Flip & Rotate / Align to Glyph Bounds); below it, full width = Copy / Paste and Drawing
- Flip and rotate run from icon buttons, with a 9-axis (3x3) widget to set the pivot; icon colors adapt to the light / dark UI
- External changes (e.g. the Preferences dialog) sync when you click (re-activate) the palette
- Press Esc to close while the palette is active

### Panels & options

- Key input: cursor step (cursorKeyLength); switch ruler unit via popup, adjust with Up/Down / Shift / Option
- Align Options: Preview Bounds
- Align to Glyph Bounds: Point Type / Area Type
- Transform Options: Pattern Tiles / Corners / Strokes & Effects
- Transform: Flip horizontal / vertical and rotate (counterclockwise / clockwise) from icon buttons. A 9-axis (3x3) anchor widget sets the pivot (center by default), computed from the selection's overall visible bounds. Icon colors adapt to the light / dark UI
- Copy / Paste: Paste without Formatting / Paste Remembers Layers
- Drawing: Real-time Drawing & Editing / Refresh Preview (GPU preview)

### Script info

- Version: v2.0.0
