# Generate a specimen sheet of installed fonts

[![Direct](https://img.shields.io/badge/Direct%20Link-FontCatalogGenerator.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FontCatalogGenerator.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontCatalogGenerator.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Lists the fonts installed on the system and generates a specimen sheet for them on the artboard.

### Features

- Sample string and font size are set in a dialog
- Enumerates the installed fonts and lays out a specimen for each

### Usage

1. Run the script.
2. Set the sample string and the font size.
3. Run it, and the specimens are generated on the artboard.

### Notes

- Systems with many fonts can take a while.
- TypefaceSampler.jsx and FontSampler.jsx cover similar ground.

### Update History

- v1.7.11 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.7.10 (2026-10-03) Units now appear inside the numeric fields; removed the unit labels next to the fields
- v1.7.9 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.7.8 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.7.7 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.7.6 (2026-09-30) : Button rows with only right-side buttons are now centered
- v1.7.5 (2026-09-30) : Fixed an error when running with characters selected by the Type tool
- v1.7.4 (2026-09-29) : Dialog opacity changed to 98%
- v1.7.3 (2026-09-29) : When the type unit is Q, the unit now reads "Q" instead of "Q/H" (unit handling now goes through the shared table)
- v1.7.2 (2026-09-28) : English field labels now end in ":" without a trailing space. The button row is now built with the shared part
- v1.7.1 (2026-09-28) : The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.7.0 (2026-09-27) : Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.6 (2026-01-27)
