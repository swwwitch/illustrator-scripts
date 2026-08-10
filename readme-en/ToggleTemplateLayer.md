# ToggleTemplateLayer

[![Direct](https://img.shields.io/badge/Direct%20Link-ToggleTemplateLayer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/ToggleTemplateLayer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ToggleTemplateLayer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Turns the active layer's template attribute (locked, non-printing, dimmed images) on or off.
- A small dialog chooses ON (make template) or OFF (release).
- Executed through a dynamic action.
- Reads the active layer name before running and applies only the attributes, without renaming.
- ON does nothing on a locked layer; OFF works even on a locked (template) layer.
- Hidden layers are excluded from both ON and OFF.

### Update History

- v1.0 (20240721): Initial version
- v1.1 (20260601): Reorganized the temporary-action creation and execution into a standard pattern
- v1.2 (20260601): Read the active layer name dynamically and inject it into parameter-3
- v1.3 (20260601): Added template OFF, with a small dialog to choose ON/OFF

### Script info

- Version: v1.3
