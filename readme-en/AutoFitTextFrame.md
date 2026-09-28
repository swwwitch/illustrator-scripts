# Apply Auto Size to area type in one click

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoFitTextFrame.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AutoFitTextFrame.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoFitTextFrame.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Applies Auto Size to the selected area type, growing each frame until the overset text fits.
- A dialog-free version of the Auto Size option in [FitAreaText](FitAreaText.md).

### How to Use

1. Select area type (one or more).
2. Run the script.

### Targets

- Area type
- Area type inside groups (groups are traversed)
- The area type you are editing with the text cursor

### Notes

- Frames only grow, and Auto Size stays on afterwards (it is not turned off again).
- Point type and type on a path are not supported.
- Locked, hidden, and non-editable text is skipped.
- Auto Size is applied through a temporary action. If it fails, the remaining frames are left as they are and an alert is shown.

### Update History

- v1.0.1 (2026-09-28) Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure; a failure now shows an alert and stops
- v1.0.0 (2026-09-26) Initial release
