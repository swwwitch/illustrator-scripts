# Turn the active layer into a template layer

[![Direct](https://img.shields.io/badge/Direct%20Link-MakeTemplateLayer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/single-function/MakeTemplateLayer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeTemplateLayer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Turns the active layer into a template layer (locked, non-printing, dimmed images). ON only.
- Runs immediately, with no dialog.
- Executed through a dynamic action.
- Reads the active layer name before running and applies only the attributes, without renaming.
- Does nothing on locked or hidden layers.

### Script info

- Version: v1.0.3

### Update History

- v1.0.1 (2026-09-27): Messages now appear in the UI language only
- v1.0.2 (2026-09-28): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure
- v1.0.3 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
