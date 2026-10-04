# Turn a group's bounds into an artboard with a margin

[![Direct](https://img.shields.io/badge/Direct%20Link-Group2Artboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/Group2Artboards.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/Group2Artboards.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script for Illustrator that automatically adds new artboards around each selected group object with an optional margin.
- Supports multiple group selections, sequential naming, file name reference, and deletion of existing artboards.

Last Updated: 2025-08-22 (v1.3)

### Main Features

- Automatically generate artboards based on group objects
- Set margin value and choose between preview or geometric bounds
- Flexible naming: prefix, symbol, sequence number, file name reference
- Option to delete existing artboards
- Japanese and English UI support

### Process Flow

1. Select group objects
2. Configure margin and artboard name options in the dialog
3. Click OK to add artboards automatically

### Update History

- v1.0 (20250703): Initial version
- v1.1 (20250704): Cleaned comments and optimized logic
- v1.2 (20250705): Adjusted behavior for clip groups
- v1.3 (20250822): Added keyboard increments (Up/Down, Shift+Up/Down, Option+Up/Down)
- v1.4.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.4.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.4.2 (20260928): The button row is now built with the shared part
- v1.4.3 (20260929): Dialog opacity changed to 98%
- v1.4.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.4.5 (20260930): Button rows with only right-side buttons are now centered
- v1.4.6 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.4.7 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.4.8 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.4.9 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

### Script info

- Version: v1.4.9
