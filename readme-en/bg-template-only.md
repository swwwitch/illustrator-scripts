# Give the target layer the template attribute

[![Direct](https://img.shields.io/badge/Direct%20Link-bg--template--only.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/single-function/bg-template-only.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/bg-template-only.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Loads and runs a temporary action that gives the target layer the template attribute (locked, non-printing, dimmed images).
- The temporary action is unloaded afterwards.

### Usage

1. Make the target layer active
2. Run the script

### Notes

- There is no dialog.
- Use ToggleTemplateLayer.jsx when you need to switch the attribute on and off.

### Update History

- v1.0.1 (2026-09-27): Error messages now appear in English on English systems
- v1.0.2 (2026-09-28): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure
- v1.0.3 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
