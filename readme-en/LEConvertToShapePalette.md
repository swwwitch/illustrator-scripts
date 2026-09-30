# Apply the Convert to Shape live effect from a palette

[![Direct](https://img.shields.io/badge/Direct%20Link-LEConvertToShapePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/single-function/LEConvertToShapePalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LEConvertToShapePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that applies the "Convert to Shape" live effect to the selection. Pick Rectangle/Ellipse and Absolute/Relative sizing plus width/height (pt); the selection updates as a live preview. All DOM work is delegated to the main engine via BridgeTalk.

### Script info

- Version: v1.1.1

### Update History

- v1.1.2 (2026-10-01) Moved the Apply button from inside the panel to the standard button row at the bottom of the palette. Unified the window and panel margins and spacing with the shared layout part
- v1.1.1 (2026-09-28) Removed the space after the colon in English field labels (shared localization helpers)
- v1.1.0 (2026-09-27) Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.0.1 (2026-09-26) Renamed the file from `LEConvertToShape.jsx` to `LEConvertToShapePalette.jsx`.
