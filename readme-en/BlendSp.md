# Create and adjust a blend

[![Direct](https://img.shields.io/badge/Direct%20Link-BlendSp.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/blend/BlendSp.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BlendSp.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Creates, configures and adjusts a blend from a single dialog, depending on what is selected.

When the selection already contains a blend it only configures and adjusts it; otherwise it creates the blend first and then configures it.

### Features

- Step count as an integer from 0 to 1000
- Step slider (0–32 normally, 0–128 with Option, 0–1000 with Shift)
- Blend orientation and Reverse
- Release / Expand / Replace Spine
- Live preview of steps and orientation (undo-based)
- Japanese / English UI

### Usage

1. Select a blend, or the objects you want to blend.
2. Run the script.
3. Adjust the step count and the orientation while watching the preview.
4. Confirm with OK.

### Notes

- When a blend is selected, its current step count is read and used as the initial value (8 if it cannot be read). The default is also 8 when no blend is selected.
- Release, Expand and Replace Spine are never run during preview; they execute only when OK is pressed.
- When the "Other" option is set to anything but None, the orientation and Reverse panels are disabled (their state is kept).
- Selecting a path inside a blend also works — the script looks up the owning blend — but results vary between environments.

### Update History

- v1.3.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.3.5 (2026-10-01): The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.3.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.3.3 (2026-09-29): Dialog opacity changed to 98%
- v1.3.2 (2026-09-28): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure
- v1.3.2 (2026-09-28): The button row is now built with the shared part
- v1.3.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2 (2026-01-01)
