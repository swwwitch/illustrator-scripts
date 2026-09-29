# Apply a kinsoku preset

[![Direct](https://img.shields.io/badge/Direct%20Link-AiSetKinsoku.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/AiSetKinsoku.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSetKinsoku.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Applies a kinsoku (line-breaking) preset to the selected text.

### Usage

1. Select the text.
2. Run the script.
3. Pick the kinsoku preset and apply it.

### Notes

- The preset order and the value passed to `paragraphAttributes.kinsoku` are set in the user settings at the top of the script.

### Update History

- v1.0
- v1.0.1 (2026-09-27): Renamed the dialog and panel titles, added English UI and tooltips
- v1.0.2 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.3 (2026-09-29): Dialog opacity changed to 98%
- v1.0.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
