# Relink missing images from a chosen folder

[![Direct](https://img.shields.io/badge/Direct%20Link-RelinkMissingLinks.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/RelinkMissingLinks.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkMissingLinks.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Detects missing linked images and relinks them automatically from a folder you choose, matching by file name.

### Features

- Finds every missing link in one pass
- The source folder is chosen in a dialog
- Matching mode selectable as exact match including extension, name only, or extension-preferred

### Usage

1. Open the document with the missing links.
2. Run the script.
3. Choose the source folder and the matching mode.

### Notes

- Links with no matching file are left untouched.

### Update History

- v1.4 (2025-08-02)
- v1.4.3 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.4.4 (2026-09-28): The button row is now built with the shared part
- v1.4.5 (2026-09-29): Dialog opacity changed to 98%
- v1.4.6 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.4.7 (2026-09-30): Button rows with only right-side buttons are now centered
