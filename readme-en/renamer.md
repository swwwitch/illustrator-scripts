# Rename artboards, layers, symbols and styles in bulk

[![Direct](https://img.shields.io/badge/Direct%20Link-renamer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/renamer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/renamer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Renames artboards, layers, symbols and graphic styles in bulk using find-and-replace and numbering.
- The target is switched in the dialog, and the resulting names are reviewed in a list before committing.

### Main Features

- Find and replace, with regular expression support
- Prefix and suffix
- Numbering, with a configurable separator and start number
- Sorting (original order / name ascending / name descending / changed first)
- Move entries within the list (↑↑ ↑ ↓ ↓↓)

### Usage

1. Run the script
2. Choose the target (artboard / layer / symbol / graphic style)
3. Set find-and-replace, prefix/suffix and numbering
4. Review the list and click OK

### Notes

- Shows an alert and stops when no document is open.

### Update History

- v1.1.7 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.1.6 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.1.5 (2026-09-30): Button rows with only right-side buttons are now centered
- v1.1.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.1.3 (2026-09-29): Dialog opacity changed to 98%
- v1.1.2 (2026-09-28): The button row is now built with the shared part
- v1.1.2 (2026-09-28): In English, removed the extra space after field-label colons and the stray Japanese counter shown after counts
- v1.1.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
