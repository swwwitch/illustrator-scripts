# SelectionToNew.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-SelectionToNew.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SelectionToNew.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionToNew.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script creates a new layer, a new artboard, or a new document from the selected objects.

What to create is chosen with radio buttons in a dialog, along with the name to give it.

## Main features

- Dialog to choose what to create (Layer / Artboard / Document)
- The new layer / artboard / document name is set in the dialog
- Keep original objects switches between moving and duplicating
- Layer: creates a new layer holding the selected objects
- Artboard: inherits the existing arrangement, with direction and spacing set in the dialog
- Document: creates a duplicate document containing only the selection
  - Locked and hidden objects can optionally be kept
- Keeps the original stacking order
- `L` / `A` / `D` keys switch the target
- Tooltips on every control
- Dialog settings persist while Illustrator is running
- Japanese and English UI

## Usage

1. Select the objects to process.
2. Run `SelectionToNew.jsx`.
3. Choose what to create (or press `L` / `A` / `D`) and enter a name.
4. Set the options as needed (artboard direction and spacing, Keep original objects, and so on).
5. Click OK to run.

## Behavior per target

| Target | Behavior |
| --- | --- |
| Layer | Creates a new layer at the top and moves (or duplicates) the selection into it. A numeric suffix such as "Name 2" is added when the name is taken. |
| Artboard | Creates a new artboard matching the active artboard's size and moves (or duplicates) the selection onto it, keeping the relative position. The existing layout (rows and columns) is analyzed so its pitch and column count are inherited. Direction and spacing are set in the dialog. |
| Document | Saves the document under a new name to build the duplicate, inverts the selection and deletes it, then removes every artboard but the current one. The source file is untouched on disk and is reopened afterwards. |

## Create panel (target and name)

Each row pairs a target radio button with the name field for that target. Only the field matching the selected target is enabled; the others are dimmed.

| Target | Default name |
| --- | --- |
| Layer | "New Layer" ("新規レイヤー" in Japanese) |
| Artboard | Active artboard name + `_new` |
| Document | Source file name + `_selection` (the source extension is shown to its right) |

- A blank field falls back to the default.
- Press `L` (Layer), `A` (Artboard) or `D` (Document) to switch targets. The shortcuts are disabled while a name field has focus, so the keys type normally.
- The layer name gets a numeric suffix such as "Name 2" when a layer with that name already exists.
- The document is created next to the source file. The source file's own name is rejected, and an existing file prompts for confirmation before it is overwritten.

## Artboard placement panel

Enabled for Artboard only.

| Setting | Description |
| --- | --- |
| Direction | Whether the new artboard goes to the Right (horizontal) or Down (vertical) of the current one. |
| Spacing | Gap between artboards, entered in the ruler unit. The default is inferred from the existing layout, falling back to the preference value when there are fewer than two artboards. |

The spacing field responds to the arrow keys (Shift for ±10 snapped to multiples of 10, Option/Alt for ±0.1).

## Options panel

| Checkbox | Enabled for | Behavior |
| --- | --- | --- |
| Keep original objects | Layer / Artboard | On keeps the originals and places copies; off (default) moves the selection. |
| Include locked objects | Document | Keeps locked objects, including art on locked layers, instead of deleting them. |
| Include hidden objects | Document | Keeps hidden objects, including art on hidden layers, instead of deleting them. |

Document always keeps the source document, so Keep original objects is dimmed for it. Conversely, Layer and Artboard have no scope to define the two Include options against, so those are dimmed for them.

Hover any control for a tooltip explaining what it does and why it may be dimmed.

When the Include options keep objects:

- Only objects **on the current artboard** are kept; anything outside it is deleted.
- Only **top-level objects** (direct children of a layer) are covered. Locked or hidden objects inside a group or compound path follow their parent: they stay if the parent stays, and go if the parent is deleted.

## User settings

The artboard behavior can be changed in the "ユーザー設定 / User settings" block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `ARTBOARD_INSERT_AFTER_CURRENT` | `true` | `true` inserts after the current artboard, `false` at the end |
| `ARTBOARD_DIRECTION_AXIS` | `0` | Initial Direction in the dialog. `0` runs right (horizontal), `1` runs down (vertical) |
| `ARTBOARD_NAME_SUFFIX` | `"_new"` | Suffix for the artboard name default |
| `DOCUMENT_NAME_SUFFIX` | `"_selection"` | Suffix for the document name default |
| `SHORTCUT_TARGETS` | `{L, A, D}` | Keys that switch the create target |

## Notes

- Shows an alert and exits when no document is open or nothing is selected.
- The artboard direction and spacing are set in the dialog. The spacing default is inferred from the existing layout, falling back to the preference value when there are fewer than two artboards.
- Artboard mode anchors on the artboard holding the selection's center, falling back to the active artboard.
- Artboard mode assumes **the selection sits on a single artboard**. A selection spanning several artboards cannot be placed correctly, because the shift is measured as one value.
- Document mode also deletes guides: they are never part of the selection, so Select > Inverse marks them for deletion.
- Artboard mode only translates the objects; their layer and group nesting is unchanged.
- When the existing layout does not run along the chosen direction, or there is only one artboard, the existing artboards are left untouched and the new one is placed relative to the artboard just before the insert point.
- Document mode requires a saved document with no unsaved changes, because those changes would be lost when the source document is reopened.
- The duplicate is written next to the source file under the name given in the dialog. **The duplicate document stays bound to that file.**
- Document mode closes and reopens the source document. Its contents on disk are unchanged, but **the source document's undo history is lost.**
- Document mode carries swatches, symbols and document settings over as-is. Because it only deletes the other objects, layers that end up empty remain, and every layer keeps its original visibility and lock state (hidden layers stay hidden).
- Document mode finds what to delete with Select > Inverse. Inverse skips locked and hidden art, so everything is unlocked and shown first; the layer states are then restored.
- Layer mode **moves** the selection into the new layer, so **selecting an object inside a group or compound path pulls it out of its parent.** That changes the group's structure and can change the appearance (stacking, knockout groups, and so on). Select the group itself to move it as a whole.
- Layer mode sorts items by layer and container stacking order (zOrderPosition) before processing. A selection that mixes grouped items with top-level items may not reproduce the original order exactly.
- The dialog settings (target, direction and options) persist only while Illustrator is running and reset on restart. Names and spacing always start from the document-specific default.
- Document mode shows a progress palette, because saving, deleting and reopening take time.

## Credits

The artboard layout logic is ported from [AddArtboardPlus.jsx](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddArtboardPlus.jsx).

- Original: Copyright (c) 2018 Takeshi Umeda (noellabo) / [dtp-discourse.jp](https://dtp-discourse.jp/t/illustrator/99)
- Largest canvas bounds: original idea by OMOTI

## Update history

- v1.0.1 (2026-09-05): Fixed the dialog settings never being remembered between runs (they are now kept on `$.global` in the persistent engine)
- v1.0.0 (2026-07-29): Initial version
