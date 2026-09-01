# FitToArtboardWidth


[![Direct](https://img.shields.io/badge/Direct%20Link-FitToArtboardWidth.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/FitToArtboardWidth.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitToArtboardWidth.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Groups the selection by the artboard each object overlaps, then scales every group as one cluster to the artboard width while keeping its aspect ratio and centers it on that artboard. By default the result is 90% of the artboard width. When fitting to the width would make the group taller than the artboard, it is scaled to fit 90% of the artboard height instead.

There is no dialog; running the script applies the change directly. It is the "Artboard → Width" base of [SmartObjectResizer](SmartObjectResizer.md), split out so it runs in one step with no options to pick.

### Features

- Handles a selection spanning several artboards artboard by artboard
- Resizes the objects on the same artboard as one cluster, keeping the aspect ratio
- Switches to a height base automatically when fitting to the width would overflow the artboard height
- Does not group anything, so the parent hierarchy and stacking order stay untouched
- Scales stroke weights, patterns and gradients by the same factor
- Picks the artboard that overlaps each object most as its reference
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
| `HEIGHT_PERCENT` | `90` | Target height as a percentage of the artboard height, used when the script switches to a height base |
| `USE_PREVIEW_BOUNDS` | `true` | Bounds used for measuring: `true` = preview bounds (incl. strokes and effects), `false` = geometric bounds (path edges) |
| `CENTER_VERTICALLY` | `true` | Set to `false` to center horizontally only and keep the vertical position |

### How it works

1. Check that a document is open and something is selected
2. If characters are selected, reselect the text object instead
3. Assign each object to its reference artboard (largest overlap, or the nearest one when nothing overlaps)
4. Handle the rest per artboard, measuring the rectangle enclosing that group
5. Compute the scale factor from artboard width × `WIDTH_PERCENT` ÷ 100
6. If that factor makes the group taller than the artboard, recompute it from artboard height × `HEIGHT_PERCENT` ÷ 100
7. Scale each object's size and its offset from the cluster's top-left by that same factor
8. Move the group to the center of the artboard

### Notes

- With multiple objects on the same artboard, the gaps between them scale by the same factor. **To fit each object to the artboard width individually, use** [SmartObjectResizer](SmartObjectResizer.md).
- When `USE_PREVIEW_BOUNDS` is `true`, measuring uses the visual edges including strokes and effects, so stroke weights scale by the same factor.
- In a multi-artboard document the reference is the artboard that overlaps each object most, not the active one. A selection spanning several artboards is resized and centered independently on each of them.
- The height base only kicks in automatically when the width base overflows. To always use a height base, or for a bleed base or one-side-only resizing, use [SmartObjectResizer](SmartObjectResizer.md).

### Update history

- v1.0.0 (20260821) : Initial release
- v1.0.1 (20260901) : Switches to a height base (90% by default) when fitting to the width would overflow the artboard height; handles a selection spanning several artboards artboard by artboard
