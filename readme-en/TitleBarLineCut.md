# Build a title bar whose rule breaks around the text

[![Direct](https://img.shields.io/badge/Direct%20Link-TitleBarLineCut.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/TitleBarLineCut.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TitleBarLineCut.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

With one text frame and one rectangle path selected, builds a title bar whose rule is cut away around the text.

### Features

- Margin, corner radius, fill, notch and stroke weight are set in a dialog
- Japanese / English UI

### Usage

1. Select one text frame and one rectangle path.
2. Run the script.
3. Set the margin and the other options, then click OK.

### Notes

- The margin must be zero or greater.
- Dialog values persist within a session and reset when Illustrator restarts.
- Use TitleBarLineCutSp.jsx when you only need the rule cut, without the bar.

### Update History

- v1.2.5 (2026-09-30) Button rows with only right-side buttons are now centered
- v1.2.4 (2026-09-30) Fixed an error when running with characters selected by the Type tool
- v1.2.3 (2026-09-29) Dialog opacity changed to 98%
- v1.2.2 (2026-09-28) The button row is now built with the shared part
- v1.2.1 (2026-09-28) The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.0 (2026-09-27) Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.1.2
