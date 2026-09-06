# Copy guides as stroked paths

[![Direct](https://img.shields.io/badge/Direct%20Link-CopyGuidesAsPaths.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/CopyGuidesAsPaths.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyGuidesAsPaths.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Converts the guides on the active artboard into ordinary stroked paths and sends them to the clipboard. The original guides stay in the document, so you can carry the guide layout into another document or application as visible rules without losing the guides themselves.

### Features

- Targets only the guides on the active artboard (matched by each guide's center point)
- Includes locked objects and guides on locked layers
- Keeps one guide out of any set stacked at the same position and deletes the extras
- Reproduces curves and closed paths, not just straight lines
- Builds the paths on a temporary layer and cuts them, so nothing is left behind in the document
- Stroke color follows the document color mode: K100 for CMYK, black for RGB

### Usage

1. Make the artboard holding the guides active.
2. Run the script.
3. When "Copied n guide(s)." appears, paste into the destination.

### How it works

- Collects the guides on the active artboard
- Deletes guides stacked at the same position
- Creates a temporary layer and rebuilds each guide from its anchors and direction handles as a stroked path
- Selects only the new paths and cuts them to the clipboard
- Removes the temporary layer and restores the previously active layer

### Options

Edit the "ユーザー設定 / User Settings" block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `STROKE_WIDTH` | `0.5` | Stroke width applied to the converted paths (pt) |
| `MATCH_TOLERANCE` | `0.001` | Tolerance for treating two guides as the same position (pt) |
| `TEMP_LAYER_NAME` | `"__guide_to_path__"` | Name of the temporary layer |

### Notes

- Hidden guides and guides on hidden layers are skipped.
- Artboard membership is decided by the guide's center point, so a long guide spanning several artboards is skipped when its center falls outside the active one.
- Removing duplicated guides modifies the document (undoable).
- With "Paste Remembers Layers" enabled, pasting creates a `__guide_to_path__` layer in the destination document.
- Snapped guides can differ by around 1e-12, so duplicate detection uses the `MATCH_TOLERANCE` threshold.

### Update history

- v1.0.0 (20260907) : Initial release
