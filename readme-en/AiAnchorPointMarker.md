# AiAnchorPointMarker

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAnchorPointMarker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/AiAnchorPointMarker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAnchorPointMarker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

When you illustrate how a path is built, you often need the anchor points to be *visible*. Illustrator draws its own little squares on screen, but those are an editing aid — they never make it into an exported image or a print. So you end up drawing one square, copying it, and dropping copies onto the anchors by hand.

This script takes over that part. It collects every anchor point of the selection and places one marker at each of them: a square of the size and color you specify, or an object of your own, or a symbol.

You decide what to place and how in a dialog, but **every change is drawn onto the real artboard right away**, so you can judge the result before you commit.

<img alt="" src="" width="50%" />

## Usage

1. Select the objects (groups and compound paths are fine as they are)
2. Run the script
3. Adjust what to place, the size, the color and the options while watching the preview
4. [OK]

If the selection contains no paths at all (only text or images), the script says so and exits.

While it runs, edges (the on-screen anchor display) and the Live Corner Annotator are hidden — otherwise the preview squares and Illustrator's own display sit on top of each other and become impossible to tell apart. Both are restored when the script finishes.

## Object to add

| Kind | What gets placed |
| --- | --- |
| Auto-generate square | A square of the given size and color |
| Frontmost object | A duplicate of the frontmost object **within the selection** |
| Symbol | An instance of a symbol in the document |

"Frontmost object" looks for the frontmost item **inside the selection**. The document is scanned front to back and the first *selected* item wins, so an unselected object that happens to sit on top is never used as the source. Hidden and locked layers are skipped as well.

The anchors of the source object itself are excluded from the placement targets. Select a rectangle plus a path and the four corners of that rectangle do not get rectangles of their own. When only one object is selected, removing the source would leave nothing to place onto, so this option is disabled.

"Symbol" is disabled when the document has no symbols.

## Anchor Point (square settings)

These apply to "Auto-generate square".

**Size** is in points. Decimals are accepted, and the arrow keys step it: ↑↓ for ±1, Shift + ↑↓ for ±10 (snapping to multiples of 10), ⌘ + ↑↓ for ±0.1.

**Color** is set through the Choose... button, with R / G / B sliders and numeric fields. The script uses its own dialog rather than the system color picker because, on some setups, opening the standard color palette on top of a modal dialog leaves the parent dialog unresponsive. The default is RGB(79, 128, 255).

**Symbolize** (on by default) registers the square as a symbol and places instances of it. Later, when you want the anchor markers in a different color, editing the symbol changes all of them at once.

A fresh symbol is registered on every run. Existing symbols are never reused, so running the script again with a different size or color leaves the markers you placed earlier untouched. If a symbol named "アンカーポイント" already exists, only the name falls back to Illustrator's default.

## Options

| Option | What it does |
| --- | --- |
| Scale | Scale (%) for the frontmost object / symbol |
| Move to layer | Moves the placed markers to the "_anchorpoint" layer (created if missing) |
| Group | Groups the placed markers into one group (on by default) |
| Registration point | Which part of the marker aligns to the anchor (3×3) |

In "Auto-generate square" mode, **Scale and the registration point are dimmed**. The size is already given in points, so the scale is always 100%, and the registration point is fixed to center — for a square centered on the anchor, no other choice is meaningful.

Pick the top-left cell and the marker's top-left corner lands on the anchor. That is what you want for arrows or labels that should sit next to an anchor rather than on it.

With both Group and Move to layer enabled, the markers are **grouped first and then moved** — individual objects never wander off to the layer one by one.

When the script finishes, the placed markers (or the group, if you grouped them) are left selected, ready to be moved or aligned.

## Live preview

The preview is drawn into a dedicated layer (`__ANCHOR_MARKER_PREVIEW__`). Every change in the dialog discards that whole layer and redraws it.

Removal goes through the **reference to the layer the script created**, not through the layer name. Even if a layer with the same name already exists in the document, nothing but the script's own layer is ever deleted. The preview is always cleared before the real placement, on OK and on Cancel alike.

Note that the preview always draws as if Symbolize were off: registering a symbol on every redraw would flood the Symbols panel. The appearance — size, color, position — is identical to the final result.

## Targets

PathItem, CompoundPathItem, GroupItem (anchor points are collected recursively)

Groups and compound paths are walked all the way down. Text and placed images hold no path points, so they contribute no anchors.

## Notes

The dialog follows the Japanese / English UI. The 3×3 registration widget is drawn by the script rather than built from radio buttons, and it reads Illustrator's UI brightness setting to switch between the light and dark palettes.

## Change log

- v1.0.1 (2026-07-06): Changed the default anchor point color to RGB(79, 128, 255)
- v1.0.0 (2026-07-05): Initial version

### note

- [【Illustrator】解説画像で使う「アンカーポイント表示」を自作オブジェクトで自動配置するスクリプト](https://note.com/dtp_tranist/n/n757f8802dc4b)
