# Apply fonts used in the document

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyDocumentFonts.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ApplyDocumentFonts.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyDocumentFonts.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script that collects fonts used in the document and displays them sorted by usage count.
- You can immediately apply a selected font to text objects and export the list as a text file.

### Main Features

- Collect and display document fonts sorted by usage count
- Filter fonts using a search filter
- Instantly apply selected font to currently selected text objects
- Export font list as a text file on the desktop
- Japanese and English UI support

### Process Flow

1. Collect font information from the document
2. Display fonts in a dialog, filterable via search
3. Apply font instantly when selected
4. Optionally export font list to a text file

### Update History

- v1.0.0 (20250225): Initial version
- v1.1.0 (20250228): Added export feature
- v1.1.1 (20250301): Supported applying to text inside groups
- v1.1.2 (20250302): Adjusted font count method for groups
- v1.1.4 (20260927): Cancel now also restores the fonts of text inside groups. The list heading now matches the actual order (by name), and field labels gained colons and tooltips
- v1.1.5 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.6 (20260928): The button row is now built with the shared part
- v1.1.6 (20260928): Target collection now uses the shared part
- v1.1.7 (20260929): Dialog opacity changed to 98%
- v1.1.8 (20260930): Fixed an error when running with characters selected by the Type tool
