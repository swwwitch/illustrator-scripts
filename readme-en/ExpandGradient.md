# Split a gradient into solid-color objects

[![Direct](https://img.shields.io/badge/Direct%20Link-ExpandGradient.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ExpandGradient.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandGradient.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Splits the gradient on the selected objects into a given number of solid-color objects.

Post-processing can tidy the overlaps into a single set, or build a blend from the two end objects.

### Usage

1. Select the object with the gradient.
2. Run the script.
3. Set the step count and the post-processing, then click OK.

### Options

**Steps**

The number of solid-color objects to produce. Defaults to 5; any integer of 2 or more. The stepper buttons and Up/Down adjust by 1; with Shift they snap to the next multiple of 10.

With "Convert to blend" selected the value is fixed at 2 and the field is dimmed.

**Post-processing**

- None: Object > Expand only (the clipping group is kept)
- Simple expand: apply Pathfinder Crop and Merge to tidy the result into solid-color objects
- Convert to blend: after the above, keep only the two end objects and build a blend

### Notes

- A temporary .aia action is generated, loaded and played internally, then unloaded and deleted. The action set and action names are kept in constants, and the names inside the .aia are generated from the same values.
- Illustrator's Expand yields one object fewer than specified, so the script compensates for it.
- The step count of "Convert to blend" comes from Illustrator's current Blend Options setting.

### Article

[Split a gradient into solid-color objects (Japanese)](https://note.com/dtp_tranist/n/nbe084e691ba5)

### Update History

- v1.2.8 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.2.7 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.2.6 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.2.5 (2026-09-30): Button rows with only right-side buttons are now centered
- v1.2.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.2.3 (2026-09-29): Dialog opacity changed to 98%
- v1.2.2 (2026-09-28): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure. The button row is now built with the shared part. Target collection now uses the shared part
- v1.2.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten). Also fixed an error when opening the dialog and the missing tooltips
- v1.1.2 (2026-08-27): Fixed the object count coming out one short; the step count is fixed at 2 and dimmed for "Convert to blend"; added the article link; removed ExpandGradient-v2.jsx; cleaned up the code
- v1.1.0 (2026-05-25)
