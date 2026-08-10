# ResizeArtboardsAll

[![Direct](https://img.shields.io/badge/Direct%20Link-ResizeArtboardsAll.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ResizeArtboardsAll.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResizeArtboardsAll.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Resize artboards to the specified width/height with live preview.
- When nothing is selected, each artboard is adjusted individually based on items inside it.

### Key Features

- Live preview (debounced app.redraw for performance)
- Target artboards: Active only / All / Specify (1-based ranges & lists, e.g., 1-3 / 1,3 / 2-4,7)
- Anchor: Top-Left / Center
- Units follow the document ruler (snap top-left to integer when in px)
- Dialog position & opacity persistence across sessions
- Arrow keys: Up/Down = ±1, Shift = snap to multiples of 10, Option(Alt) = ±0.1 (rounded to integer at commit)

### Process Flow

1. Enter width/height (optionally choose target & anchor)
2. See instant preview of artboard updates
3. Press OK to commit, Cancel to restore

### Update History

- v1.0 (2025-08-29): Initial release

### Script info

- Version: v1.0
