# Batch move objects to a chosen layer

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartLayerManage.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/SmartLayerManage.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartLayerManage.md)


[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script to batch move objects to a specified layer.
- Supports switching modes: selected objects, all text, all objects, or all (force).

### Main Features

- Mode selection (Selected / All Text / All / All (Force))
- Option to delete empty layers (excluding layers starting with bg or //)
- Automatically change target layer color to RGB(79,128,255)
- Unlocking, showing, and recursive item collection
- Japanese and English UI support

### Process Flow

1. Select mode and target layer in the dialog
2. Collect target objects
3. Move objects to the selected layer
4. Optionally delete empty layers
5. Change target layer color

### Update History

- v1.0.0 (20250703): Initial release
- v1.0.1 (20250703): Added layer color change function
- v1.0.2 (20250703): Improved auto selection detection and empty layer deletion logic
- v1.0.3 (20250704): Added "All (Force)" mode (merge all layers)
- v1.0.6 (20260927): Dialog title changed to "Move Objects to Layer"; the All / All (Force) tooltips now describe what they actually do; "Text Only" renamed to "All Text"; all messages localized
- v1.0.7 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.8 (20260928): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure
- v1.0.9 (20260929): Dialog opacity changed to 98%
- v1.0.10 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.0.11 (2026-10-01): Moved the Move/Close buttons from a right-hand column to the standard button row at the bottom. Unified the window and panel margins and spacing with the shared layout part
- v1.0.12 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.0.13 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)