# Clear appearance, keep fill and stroke

[![Direct](https://img.shields.io/badge/Direct%20Link-ClearAppearance.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ClearAppearance.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClearAppearance.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Runs Clear Appearance on the selected objects.

A dialog first asks whether to restore anything, and which restore mode fits the selection.

### Features

- Paths and compound paths get their original fill, stroke and stroke weight reapplied
- Restore mode selectable as "Fill, Stroke and Weight" or "Fill, Stroke + Stroke Details"
- The latter also restores caps, joins, dashes, dash offset and miter limit
- Opacity, blending mode and overprint can each be restored independently

### Usage

1. Select the objects whose appearance you want to clear.
2. Run the script.
3. Choose the restore mode and click OK.

### Article

https://note.com/dtp_tranist/n/na4c70c5acd60

### Update History

- v1.0.9 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.0.8 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.0.7 (2026-10-01): The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.0.6 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.0.5 (2026-09-29): Dialog opacity changed to 98%
- v1.0.4 (2026-09-28): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure; in Japanese, failure counts and detail lines now use a full-width colon. The button row is now built with the shared part
- v1.0.3 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.1 (2026-09-18)
- v1.0 (2026-04-14)
