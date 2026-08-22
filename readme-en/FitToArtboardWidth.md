# FitToArtboardWidth


[![Direct](https://img.shields.io/badge/Direct%20Link-FitToArtboardWidth.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/FitToArtboardWidth.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitToArtboardWidth.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Scales the whole selection as one cluster to the artboard width while keeping its aspect ratio, then centers it on the artboard. By default the result is 90% of the artboard width.

There is no dialog; running the script applies the change directly. It is the "Artboard → Width" base of [SmartObjectResizer](SmartObjectResizer.md), split out so it runs in one step with no options to pick.

### Features

- Resizes the whole selection as one cluster, keeping the aspect ratio
- Does not group anything, so the parent hierarchy and stacking order stay untouched
- Scales stroke weights, patterns and gradients by the same factor
- Picks the artboard that overlaps the selection most as the reference
- Centers the selection on the artboard afterwards
- Promotes a text selection made with the Type tool to the text object itself

### How to use

1. Select the objects to resize (one or more)
2. Run the script

### Settings

Edit them in the User Settings block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `WIDTH_PERCENT` | `90` | Target width as a percentage of the artboard width (`100` = full artboard width) |
| `USE_PREVIEW_BOUNDS` | `true` | Bounds used for measuring: `true` = preview bounds (incl. strokes and effects), `false` = geometric bounds (path edges) |
| `CENTER_VERTICALLY` | `true` | Set to `false` to center horizontally only and keep the vertical position |

### How it works

1. Check that a document is open and something is selected
2. If characters are selected, reselect the text object instead
3. Measure the rectangle enclosing the whole selection
4. Pick the reference artboard (largest overlap with the selection, or the nearest one when nothing overlaps)
5. Compute the scale factor from artboard width × `WIDTH_PERCENT` ÷ 100
6. Scale each object's size and its offset from the cluster's top-left by that same factor
7. Move the selection to the center of the artboard

### Notes

- With multiple objects, the gaps between them scale by the same factor. **To fit each object to the artboard width individually, use** [SmartObjectResizer](SmartObjectResizer.md).
- When `USE_PREVIEW_BOUNDS` is `true`, measuring uses the visual edges including strokes and effects, so stroke weights scale by the same factor.
- In a multi-artboard document the reference is the artboard that overlaps the selection most, not the active one.
- For a height base, a bleed base, or one-side-only resizing, use [SmartObjectResizer](SmartObjectResizer.md).

### Update history

- v1.0.0 (20260821) : Initial release
