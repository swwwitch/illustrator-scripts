# Convert point text, path text or shape plus text into area type

[![Direct](https://img.shields.io/badge/Direct%20Link-TextWithShapeToAreaType.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextWithShapeToAreaType.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextWithShapeToAreaType.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Converts point text, text on a path, or a shape plus text into area text while preserving the appearance.

### Usage

1. Select the text, and the shape if you want to use one as the frame.
2. Run the script.
3. Set the options and apply.

### Notes

- TextWithShapeToAreaTypeSimple.jsx is a lighter variant that handles point text and text on a path only.

### Update History

- v1.3.4 (2026-09-30) Fixed an error when running with characters selected by the Type tool
- v1.3.3 (2026-09-29) Dialog opacity changed to 98%
- v1.3.2 (2026-09-28) Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure. The button row is now built with the shared part. Settings are now saved through the shared part (stored in Folder.userData/illustrator-scripts/TextWithShapeToAreaTypeStyles.json)
- v1.3.1 (2026-09-28) The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.0 (2026-09-27) Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.0 (2026-07-01)
