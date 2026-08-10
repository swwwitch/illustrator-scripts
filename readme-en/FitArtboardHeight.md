# FitArtboardHeight

[![Direct](https://img.shields.io/badge/Direct%20Link-FitArtboardHeight.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/FitArtboardHeight.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitArtboardHeight.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Script name (EN)

Fit Artboard Height to Selection (Same Width)

### Overview

- Adjust the **height only** of the active artboard to the selection’s bounds plus **vertical margin** (width/left/right are preserved).
- **When nothing is selected,** adjust the **height of all artboards** individually based on items inside each artboard.

### Key features

- Live preview (uses `visibleBounds` for speed)
- Throttled `app.redraw()` during preview
- Preview-bounds toggle (apply `visible`/`geometric` only on commit)
- Target panel:
  · “Active artboard only” = with selection → active AB, without selection → all AB (legacy behavior)
  · “All artboards” = ignore selection and process all AB
- Temporary outlining for text to measure exact bounds (restored after)
- Unit label next to input, dialog position remembered (namespaced per version)

### Process Flow

1) Enter vertical margin in the dialog
2) Preview updates top/bottom only (left/right preserved)
3) Confirm to commit or cancel to restore

### Notes

- Preview always uses `visibleBounds` for performance; final commit respects the chosen `visible/geometric` option.

### Update History

- v1.2 (2025-08-25): Added Target panel (All/Active). Preview fixed to `visibleBounds` with throttled redraw. Unit label & localization cleanup. Version-scoped dialog position key.
- v1.1 (2025-08-25): Simplified to height-only adjustment. Fixed undefined `doc`. Cleaned comments and docs.
- v1.0 (2025-08-25): Initial version.

### Script info

- Version: v1.2
