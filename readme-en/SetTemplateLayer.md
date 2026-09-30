# Turn a chosen layer into a template and prefix its name

[![Direct](https://img.shields.io/badge/Direct%20Link-SetTemplateLayer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/SetTemplateLayer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SetTemplateLayer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Turns a chosen layer into a template layer and prefixes its name.

### Usage

1. Open the target document.
2. Run the script.
3. Choose the target layer and run it.

### Notes

- The layer used by the "Specified" option is set by `specifiedLayerName` ("下絵" by default).
- The prefix added to the layer name is set by `COMMENT_PREFIX` (`// ` by default).

### Update History

- v1.0
- v1.0.2 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.3 (2026-09-28): The button row is now built with the shared part
- v1.0.4 (2026-09-29): Dialog opacity changed to 98%
- v1.0.5 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.0.6 (2026-10-01) The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.0.7 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
