# Align the selection's edges or center to a direction set in code

[![Direct](https://img.shields.io/badge/Direct%20Link-GroupEdgeAlignNoFileName.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/single-function/GroupEdgeAlignNoFileName.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlignNoFileName.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Aligns the edges or the center of the selected objects in the direction given by `ALIGNMENT_SIDE`. This is GroupEdgeAlign.jsx without the filename-based direction detection.

### Features

- Direction can be `left`, `right`, `top`, `bottom`, `CENTER_X`, `CENTER_Y` or `CENTER` (default `right`)
- Snaps to a matching guide when `USE_GUIDES` is true

### Usage

1. Select the objects to align.
2. Run the script.

### Notes

- The center options never use guides; they always align to the center of the artboard.

### Update History

- v1.0 (2025-04-06)
- v1.0.2 (2026-09-27) : Alert messages are now localized for English. An invalid GUIDE_SEARCH_MODE is reported even when the document has no guides
- v1.0.3 (2026-09-29) : Clip groups are measured by their masks, now also when nested inside a group (hidden parts are left out)
