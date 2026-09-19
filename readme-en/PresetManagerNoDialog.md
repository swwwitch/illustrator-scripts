# Apply a fixed set of preferences without a dialog

[![Direct](https://img.shields.io/badge/Direct%20Link-PresetManagerNoDialog.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PresetManagerNoDialog.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerNoDialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Applies a fixed set of Illustrator preferences at once, without showing a dialog.
`ACTIVE_PRESET` at the top of the script selects which preset gets written.

### Presets

| Value | Contents |
| --- | --- |
| `minimal` | The short list: just the staples. 19 items |
| `full` | A full sweep of the preference panels. 31 items, including black appearance and ruler units, which need a restart |
| `preset1` | Same contents as [Preset 1] in PresetManager. 38 items, including guides, smart guides and the artboard highlight |

### Usage

1. Set `ACTIVE_PRESET` in the script to `"minimal"`, `"full"` or `"preset1"`.
2. Run the script.

### Notes

- The values live in `PRESET_STATES` inside the script; edit the script to change them.
- A field left out of a preset is not written, so Illustrator keeps its current value for that preference.
- Preference keys are collected in the `PREFERENCE_BINDINGS` table; adding an item means adding one row there.
- Some keys may be ignored depending on the Illustrator version; missing keys are skipped safely.
- Ruler units, UI brightness and the Shape Builder "pick color from" setting only take effect after restarting Illustrator.

### Update History

- v1.0 (2026-09-19) Merged PresetManagerNoDialogFull and PresetManagerPreset1; presets are now selected with `ACTIVE_PRESET`
- v1.0
