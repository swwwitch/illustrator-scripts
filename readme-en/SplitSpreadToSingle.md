# Split spreads into left and right single pages

[![Direct](https://img.shields.io/badge/Direct%20Link-SplitSpreadToSingle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/SplitSpreadToSingle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitSpreadToSingle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Detects spread-like objects and splits them into left and right single pages.

### Features

- Process either only the selected objects or every matching object in the document
- Supports both left-binding and right-binding page orders
- Optional sequential renaming of artboards
- Optional rearranging of artboards

### Usage

1. Select the objects you want to split (no selection is needed when processing the whole document).
2. Run the script.
3. Choose the scope and the binding direction, enable renaming and rearranging if you want them, and run.

### Notes

- Whether an object is a spread is decided from its aspect ratio.
- The binding direction determines the page order after the split.

### Update History

- v1.2 (2026-03-21)
- v1.3.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.3.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.2 (2026-09-28): The button row is now built with the shared part
- v1.3.3 (2026-09-29): Dialog opacity changed to 98%
- v1.3.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.3.5 (2026-10-01) The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
