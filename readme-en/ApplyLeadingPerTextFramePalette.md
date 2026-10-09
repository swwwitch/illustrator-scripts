# Recalculate leading from the size of each line's first characters

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyLeadingPerTextFramePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ApplyLeadingPerTextFramePalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyLeadingPerTextFramePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Recalculates the leading of each line in the selected text frames from the font size of the first few characters, and applies it.

The leading percentage is chosen in a dialog.

### Usage

1. Select the text frames.
2. Run the script.
3. Set the leading percentage and run it.

### Notes

- A partial selection (a TextRange) is normalized to its parent text frame.
- The adjustment changes the auto-leading percentage rather than the leading value itself.
- ApplyLeadingPerTextFrame110.jsx, 150.jsx and AUTO.jsx are fixed-percentage variants.

### Update History

- v1.2.4 (2026-10-10) Fixed the palette operating on another Illustrator version when several versions are running
- v1.2.3 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.2.2 (2026-10-03) Moved to the shared persistent engine `SwwwitchPalettes` so CloseAllPalettes can close it
- v1.2.1 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.2.0 (2026-09-27) Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten).
- v1.1.2 (2026-09-26) Renamed the file from `ApplyLeadingPerTextFrame.jsx` to `ApplyLeadingPerTextFramePalette.jsx`.
- v1.1.0 (2026-07-08)
