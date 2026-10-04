# Apply the Feather live effect

[![Direct](https://img.shields.io/badge/Direct%20Link-LEFeather.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/single-function/LEFeather.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LEFeather.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Applies the Effect > Stylize > Feather live effect to the current selection.

### Usage

1. Select the objects you want to feather.
2. Run the script.
3. Enter the radius and click OK.

### Notes

- The radius is shown and entered in the current ruler unit (`rulerType`) and converted to points internally.

### Update History

- v1.0.0
- v1.1.0 (2026-09-27): Added stepper buttons to the radius field. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten; Option by 0.1)
- v1.1.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.2 (2026-09-28): Fixed the radius unit label (H instead of Q for the ha ruler unit; feet, meters and yards were treated as points). Inches now shown as "in". The button row is now built with the shared part
- v1.1.3 (2026-09-29): Dialog opacity changed to 98%
- v1.1.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.1.5 (2026-09-30): Button rows with only right-side buttons are now centered
- v1.1.6 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.1.7 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.1.8 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
