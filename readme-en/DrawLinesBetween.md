# Draw horizontal rules between objects

[![Direct](https://img.shields.io/badge/Direct%20Link-DrawLinesBetween.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/DrawLinesBetween.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawLinesBetween.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Sorts the selected objects (shapes or text) from top to bottom and draws a horizontal rule between each pair.

### Features

- Input units follow the "strokeUnits" preference, both for the label and for the internal conversion to points
- Extend stretches or shrinks the rules horizontally (positive extends, negative shortens)

### Usage

1. Select the objects you want rules between.
2. Run the script.
3. Set the stroke weight and the extension, then run it.

### Update History

- v1.1.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.1.3 (20260929): Dialog opacity changed to 98%
- v1.1.2 (20260928): The button row is now built with the shared part. Fixed the line weight, extension and cap never being remembered (the script used a Photoshop-only API that Illustrator lacks; now stored in `~/Library/Application Support/illustrator-scripts/DrawLinesBetween.json`)
- v1.1.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.0
