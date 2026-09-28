# Build a white outline behind the selection

[![Direct](https://img.shields.io/badge/Direct%20Link-AddOutlineOffsetPath.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/AddOutlineOffsetPath.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddOutlineOffsetPath.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Duplicates the selection, sends it behind, applies Offset Path (a live effect), outlines it, unites it and expands it
- Groups the original with the result, runs Subtract, and fills the outcome with white

### Main Features

- Works with multiple selections
- Unit aware (pt, mm, in, cm, and so on)
- The corner join (miter, round, bevel) is configurable
- The offset value is entered in a dialog
- Dialog position and opacity can be adjusted
- The stepper buttons and arrow keys step the numeric field to the next whole number (1.5 → 2; Shift to the next multiple of 10, Option by 0.1)

### Update History

- v1.1.2 (2026-09-27): Fixed the value being re-rounded when keys other than Up/Down were pressed; added an alert when no document is open
- v1.2.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%

### Script info

- Version: v1.2.1
