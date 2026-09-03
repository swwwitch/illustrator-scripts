# Draw a rectangle the size of the artboard

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartDrawArtboardRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/SmartDrawArtboardRectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDrawArtboardRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Draws a rectangle the size of the active artboard, or of every artboard, taking an offset into account. Color, placement and target scope are set with a live preview in the dialog and committed with OK. The result can then be converted to guides or to a live shape.

### Key Features

- Offset (a negative value shrinks inward). Up/Down keys step the value (Shift = 10, Option = 0.1)
- Bleed presets (mm = 3, Q/H = 12, pt = 9; every other unit shows the 3 mm equivalent converted)
- Color modes (None / K100 at 15% opacity / HEX / CMYK)
- Placement (Front / Back / bg layer; defaults to Front)
- Target scope (Current artboard / All artboards)
- Post-draw options (Convert to Guides: default OFF / Convert to Live Shape: default ON)
- Shows the center widget (center point) when the rectangle is converted to a live shape
- Live preview (drawn on a dedicated `_preview` layer with a 1 pt dashed 50% gray stroke)
- Hotkeys (F / B / L / C / A / G)
- Button to toggle Outline and Preview display modes
- Japanese / English UI

### Usage

1. Open a document.
2. Run the script.
3. Set the offset, color, placement and target scope, check the preview, and click OK.

### Options

**Offset**

| Item | Description |
| --- | --- |
| Offset | How far the rectangle expands outward from the artboard bounds. A negative value shrinks it inward. Uses the document's ruler unit |
| Bleed | Fills in the bleed-equivalent value for the current unit. The offset field is disabled while it is ON, and reverts to your last manual value when turned OFF |

**Color**

| Item | Description |
| --- | --- |
| None | No fill and no stroke |
| K100, Opacity 15% | Black matching the document color space, at 15% opacity |
| HEX | `#RRGGBB`. Also accepts the `#RGB`, `#RG` and `#R` shorthands, color names (`red`, `orange`, …) and `gray0`–`gray100` |
| CMYK | 0–100 per channel. Empty fields are treated as 0 |

**Placement**

| Item | Description |
| --- | --- |
| Front | Brings the rectangle to the front of the target layer |
| Back | Sends the rectangle to the back of the target layer |
| bg Layer | Creates (or reuses) a `bg` layer, moves it to the back of the layer stack, and draws there |

**Target**

| Item | Description |
| --- | --- |
| Current Artboard | The active artboard only |
| All Artboards | Every artboard (disabled when the document has only one) |

**Options**

| Item | Description |
| --- | --- |
| Convert to Guides | Turns the drawn rectangles into guides and renames them to `<Guide>` (default OFF) |
| Convert to Live Shape | Converts them with Illustrator's menu command and shows the center widget (default ON) |

**Hotkeys**

| Key | Description |
| --- | --- |
| F / B / L | Front / Back / bg layer |
| C / A | Current artboard / All artboards |
| G | Toggles Convert to Guides |

Hotkeys are ignored while an edit field has focus.

### Notes

- The preview is drawn on a temporary `_preview` layer, which is removed on both OK and Cancel.
- Rectangles are drawn on the active layer when it is editable, otherwise on the first editable layer. Template layers, locked layers and the `_preview` layer are skipped; if no editable layer is found, an `_auto_draw` layer is created.
- The center widget cannot be set through the API, so it is applied by writing a recorded action to a temp file and playing it back (action set name `SmartDrawArtboardRectangle`).
- Convert to Guides and Convert to Live Shape can be used together (the conversion to a live shape happens first).

### Article

https://note.com/dtp_tranist/n/n1ba88513a9c8

### Update History

- v1.5.5 (2026-09-03) Reliable preview-layer cleanup; guide mode renames items to `<Guide>` and clears the selection; the HEX field accepts color names, shorthand hex and grayNN; the center widget is applied only when converting to a live shape. Alongside an internal cleanup (naming, structure, function splits), fixed the preview staying faint after switching to None or entering an invalid HEX value, and the mismatch between the shown offset and the drawn offset when Bleed was used with a unit other than mm, Q/H or pt
- v1.5.4 (2026-06-01) Fixed Error 8705 when the front-most layer is a template or locked layer
- v1.5.3 (2026-05-31) Renamed the object to `<Rectangle>`, tweaked the offset field width
- v1.5.2 (2026-05-31) Show the center widget, unified unit table, CMYK input fix, hotkey rework
- v1.5.1 (2025-08-24) Added post-draw options (Convert to Live Shape, Convert to Guides)
- Initial release (2025-08-20)
