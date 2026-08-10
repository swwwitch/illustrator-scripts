# QuickTransformPalette

[![Direct](https://img.shields.io/badge/Direct%20Link-QuickTransformPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/QuickTransformPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/QuickTransformPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that moves/duplicates and flips/rotates the selection immediately on an icon click (no preview, no Apply button).

- Move / duplicate: click an Up / Left / Right / Down arrow icon to move in that direction; Option-click to duplicate. The center button of the cross does not move — it duplicates the selection in place at the same coordinates (the copies become the new selection)
- Flip & rotate: icon buttons flip horizontally/vertically and rotate 90° (CCW/CW), applied immediately; Option-click duplicates before transforming; a slider at the bottom of the panel (-180 to 180°, 15° steps / Shift+drag for 90° steps) rotates by an arbitrary angle, with a numeric readout to its left showing the current angle (returns to 0° on release); a 9-axis (3x3) widget to the right of the icons sets the pivot
- Options (Margin / Preview bounds; applied to both move-duplicate and flip-rotate)
	- Margin: extra gap added after the transform, away from the anchor (unit follows the ruler); ignored for flip/rotate when the 9-axis anchor is centered
	- Preview bounds: uses the visible bounds (including stroke/effects) for size and pivot
- Keyboard: Esc closes the palette; the Margin field steps with ↑↓ (Shift = ±10, Option = ±0.1). Directional move/duplicate is by icon click only

DOM work (selection, move, duplicate, flip, rotate) is delegated to the main engine via BridgeTalk.

### Script info

- Version: v1.3.1
