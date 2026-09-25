# Apply a fixed set of preferences without a dialog

[![Direct](https://img.shields.io/badge/Direct%20Link-PresetManagerNoDialog.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PresetManagerNoDialog.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerNoDialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Applies the same preferences as [Preset 1] in PresetManager at once, without showing a dialog.

### Usage

Run the script.

### Notes

- The values live in `PREFERENCES` inside the script, one preference key and value per row; edit the script to change them.
- Some keys may be ignored depending on the Illustrator version; missing keys are skipped safely.

### Update History

- v1.1.0 (2026-09-25) Narrowed to [Preset 1] only: dropped `minimal` / `full` and the `ACTIVE_PRESET` switch, and folded the keys and values into a single table
- v1.0.1 (2026-09-25) `preset1` now matches [Preset 1] in PresetManager exactly: added Unlock on Canvas and dropped the Smart Guides / snap-to-grid settings
- v1.0 (2026-09-19) Merged PresetManagerNoDialogFull and PresetManagerPreset1; presets are now selected with `ACTIVE_PRESET`
- v1.0
