# Set the units and the keyboard increment

[![Direct](https://img.shields.io/badge/Direct%20Link-PreferenceManager--print--pt.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/single-function/PreferenceManager-print-pt.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManager-print-pt.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Sets the units and the numeric increment from a dialog.

### Features

- Reads the current unit settings and uses them as the initial values
- Changes the units and the numeric increment together
- Uses a version-specific internal mapping, so no GUI workaround is needed

### Usage

1. Run the script.
2. Choose the units and the numeric increment, then click OK.

### Notes

- Works directly, without AppleScript or Keyboard Maestro.

### Update History

- v1.0 (2025-08-06)
- v1.1.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.1.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.2 (2026-09-28): Removed the space after the colon in English field labels, matching the Japanese labels
- v1.1.2 (2026-09-28): The button row is now built with the shared part
- v1.1.3 (2026-09-29): Dialog opacity changed to 98%
- v1.1.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.1.5 (2026-10-01) The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.1.6 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.1.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
