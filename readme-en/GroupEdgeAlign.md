# Align multiple objects to the artboard edges, center, or guides

[![Direct](https://img.shields.io/badge/Direct%20Link-GroupEdgeAlign.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/GroupEdgeAlign.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlign.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Treats the selection as a single unit and aligns it to the edge or the center of the active artboard, or to a matching guide.
- No grouping is needed: the objects keep their relative positions inside the selection.
- The target is picked from a 3x3 grid of nine points. A live preview shows the result while the dialog is open, and Cancel puts everything back.

### Features

- Pick the target on the 3x3 widget in the Alignment panel (top-left / top-center / top-right / middle-left / center / middle-right / bottom-left / bottom-center / bottom-right); it sets the horizontal and vertical target at once
- Step one target at a time with the arrow keys. Each press behaves like running the script once, so with several guides in the way each press advances to the next one
- With Use guides on, guides inside the active artboard become alignment targets. Vertical guides serve left/right, horizontal guides serve top/bottom
- Use preview bounds switches between the visual bounds including strokes and effects (on) and the geometric bounds of the shape itself (off). It starts from Illustrator's Use Preview Bounds preference
- With Preview on, the result is applied on screen every time the target changes
- A clipping group is measured by the geometric bounds of its mask (the first item)
- The target can also be derived from the filename (`GroupEdgeAlignRIGHT.jsx` aligns right)
- Keyboard shortcuts
  - Target: W=top-left / E=top-center / R=top-right / S=middle-left / D=center / F=middle-right / X=bottom-left / C=bottom-center / V=bottom-right
  - G toggles Use guides, B toggles Use preview bounds
- Automatic Japanese / English switching

### Usage

1. Select the objects to align.
2. Run the script.
3. Click a cell on the 3x3 widget (or press its shortcut key).
4. Click OK.

If you only used the arrow keys, clicking OK without picking a target keeps the position they produced. Cancel restores the position from before the run, arrow-key steps included.

### Options

The defaults live in the User Settings block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `SHOW_DIALOG` | `true` | Set to `false` to skip the dialog and align immediately using the target derived from the filename |
| `USE_GUIDES` | `true` | Initial state of Use guides |
| `DEFAULT_ALIGNMENT_SIDE` | `"right"` | Target used when the filename does not name one |
| `GUIDE_SEARCH_MODE` | `"inside"` | `"inside"` takes the closest guide ahead of the selection, `"nearest"` takes the closest guide in either direction |
| `GUIDE_ORIENTATION_TOLERANCE` | `0.01` | Tolerance for treating a guide as horizontal or vertical |

### Notes

- Only guides inside the active artboard are considered.
- Picking a target on the 3x3 widget ignores guides and aligns to the artboard edge or center (Use guides goes dim). Guides apply to the arrow-key steps and to runs without the dialog (`SHOW_DIALOG = false`).
- Center targets (top-center / middle-left / center / middle-right / bottom-center) never snap to guides.
- For one-shot scripts fixed to a single direction, use the seven files under `GroupEdgeAlign-7/` (LEFT / RIGHT / TOP / BOTTOM / CENTER / CENTERX / CENTERY).
- `GroupEdgeAlignNoFileName.jsx` is the variant without filename-based detection.

### Article

[DTP Transit (Japanese)](https://note.com/dtp_tranist/n/n4ae0e1e70481)

### Update History

- v1.0 (2025-04-06) : Initial release
- v1.0.1 (2026-08-31) : Replaced the 3x3 radio buttons with an onDraw nine-point widget and reorganized the button row into left/right groups. Also trimmed the header to an overview plus a README pointer, added the article URL to the basic info block, split the user settings and layout blocks, restructured LABELS into nested categories, added JSDoc to every function, folded the per-direction if-chains into lookup tables, merged the duplicated guide-search checks, and moved the preview and arrow-key stepping into an alignment session
