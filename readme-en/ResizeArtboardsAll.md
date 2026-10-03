# Resize artboards to a given width and height

[![Direct](https://img.shields.io/badge/Direct%20Link-ResizeArtboardsAll.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ResizeArtboardsAll.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResizeArtboardsAll.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Resize artboards to the specified width/height with live preview.
- When nothing is selected, each artboard is adjusted individually based on items inside it.

### Key Features

- Live preview (debounced app.redraw for performance)
- Target artboards: Active only / All / Specify (1-based ranges & lists, e.g., 1-3 / 1,3 / 2-4,7)
- Anchor: Top-Left / Center
- Units follow the document ruler (snap top-left to integer when in px)
- Dialog position & opacity persistence across sessions
- Stepper buttons and arrow keys: step to the next whole number (1.5 → 2), Shift = snap to the next multiple of 10, Option(Alt) = ±0.1

### Process Flow

1. Enter width/height (optionally choose target & anchor)
2. See instant preview of artboard updates
3. Press OK to commit, Cancel to restore

### Update History

- v1.1.7 (2026-10-03): The number fields now show their unit inside the field; removed the unit from the panel title
- v1.1.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.1.5 (2026-10-01): The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.1.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.1.3 (2026-09-29): Dialog opacity changed to 98%
- v1.1.2 (2026-09-28): The button row is now built with the shared part
- v1.1.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.0.2 (2026-09-27): Fixed the artboard-number field staying enabled after switching back from Specify, and the width/height fields being overwritten with the active artboard's size while typing. Each preview now starts from the original sizes, so artboards dropped from the target go back. Units now come from the ruler-unit preference, including H. Tidied wording and tooltips
- v1.0 (2025-08-29): Initial release

### Script info

- Version: v1.1.7
