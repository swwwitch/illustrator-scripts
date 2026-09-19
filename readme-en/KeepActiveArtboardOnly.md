# Keep only the active artboard

[![Direct](https://img.shields.io/badge/Direct%20Link-KeepActiveArtboardOnly.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/KeepActiveArtboardOnly.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KeepActiveArtboardOnly.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Keeps only the active artboard, or removes every empty artboard in one pass.

### Features

- Removes every artboard but the active one, along with the objects and object guides outside it (turn the option off to remove the artboards only)
- Temporarily relaxes locked, hidden and template layers and groups, then restores them afterwards
- Guides are tested against the path itself, not just the bounding box
- Removes empty artboards together, with a toggle for whether hidden layers and objects count as occupying an artboard
- The target count updates as soon as the toggle changes
- Japanese / English UI

### Usage

1. Make the artboard you want to keep active.
2. Run the script.
3. Pick which artboards to remove, check the options and click OK.

### Notes

- The changes cannot be undone, so duplicating the file before running it is recommended.
- Ruler guides are never removed.
- Visibility is resolved through the ancestors, so hidden items inside groups, and items under hidden groups or layers, are treated the same way.
- "Empty artboards" is unavailable when the document has only one artboard.

### Update History

- v1.2.1: Absorbed RemoveOtherArtboards.jsx and RemoveEmptyArtboards.jsx; the artboards to remove are now chosen in a dialog
- v1.2: Initial version
