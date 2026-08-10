# DrawRectangleBehindSelectedObject

[![Direct](https://img.shields.io/badge/Direct%20Link-DrawRectangleBehindSelectedObject.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/DrawRectangleBehindSelectedObject.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawRectangleBehindSelectedObject.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Generate rectangles offset from the bounding box of selected objects
- Live preview with immediate feedback; created rectangles are always sent to back
- Opacity applies to both preview and the finalized rectangle

Last updated: 2025-11-09

### Key Features

- Offset (follows current ruler units)
- Corner radius (applied via Live Effect, kept unexpanded)
- Fill/Stroke color options (K100 / White / HEX / CMYK)
- Target: Individual or as Group
- Preview on a dedicated layer (does not pollute history)
- Dialog position, opacity, and parameter persistence

### Processing Flow

1. Compute bounding box of target objects
2. Apply offset, corner radius, and fill/stroke settings to rectangle
3. Render preview to dedicated layer; finalize on OK
4. Optionally group rectangle with original text

### Update History

- v1.0 (2025-08-22): Initial version
- v1.1 (2025-08-23): Added preview and color selection
- v1.2 (2025-08-23): Added type (fill/stroke), stroke width, and preset saving
- v1.3 (2025-08-28): Added dialog position, opacity, and parameter persistence
- v1.4 (2025-09-02):
- v1.5 (2025-11-09): Reviewed fill logic (HEX→CMYK when needed, disable overprint, enforce Normal)
- v1.6 (2025-11-09): Preview stabilization (debounce & cancel, before/afterRender, bump compat, immediate refresh fix)

### Script info

- Version: v1.6
