# Resize objects

[![Direct Link](https://img.shields.io/badge/Direct%20Link-SmartObjectResizer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartObjectResizer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectResizer.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Sometimes you just want a set of mismatched objects to line up: icons at the same size, logos matched to the largest one, photos stretched to fill the artboard.

But **what you match them to depends on the job**. Width or height? The long side? The area? And matching sizes is rarely the end of it — alignment usually follows. That means typing numbers into the Transform panel, switching to the Align panel, clicking again, and going back and forth.

This script puts the resize basis and the alignment in a single dialog, so **every radio button you click shows you the result right away**.

<img alt="" src="" width="50%" />

## Usage

1. Select two or more objects you want to match in size.
2. Run the script.
3. Click a radio button under "Resize base".

**No base is selected when the dialog opens.** Simply opening it does nothing. Resizing runs — and is previewed — only once you click a base radio.

Every time you switch the base, the selection is restored to its original state before being recalculated, so you can try several options and compare them. Click OK to keep the result, or Cancel to revert.

## Resize base

| Base | Options | Behavior |
| --- | --- | --- |
| Max | Width / Height | Matches everything to the largest width (height) in the selection |
| Min | Width / Height | Matches everything to the smallest width (height) in the selection |
| Fixed Size | Width / Height | Matches everything to the value you type |
| Ref. side | Long side / Short side | Matches each object's long (short) side to the others |
| Area | Max / Min | Matches every area to the largest (smallest) one |
| Artboard | Width / Height | Fits the whole selection to the artboard width (height) and centers it |
| Bleed | Width / Height | Fits to the artboard plus bleed (3 mm per side) and centers it |

Only one base can be checked at a time. Clicking one turns all the others off.

The "Fixed Size" field uses the document's ruler units, and starts at the average width of the selection. You can nudge the value with the arrow keys.

| Key | Step |
| --- | --- |
| ↑↓ | ±1 |
| Shift + ↑↓ | ±10 (snaps to multiples of 10) |
| Option + ↑↓ | ±0.1 |

### "Ref. side" vs. "Max / Min"

The difference between these two is the easy one to miss.

"Max (Width)" looks at **a fixed side — the width**. Landscape or portrait, it does not matter: every object ends up the same width.

"Ref. side (Long side)" looks at **whichever side is longer for that object**. Landscape objects are driven by their width, portrait ones by their height, so the result is that the *longest edge* matches across the selection. Use this when you want objects of different orientations to feel like the same visual weight.

### Artboard and Bleed are the only "as a group" bases

Where the other bases resize **each object individually**, "Artboard" and "Bleed" treat **the whole selection as one cluster**. The bounding box of the selection is uniformly scaled to the artboard width (height), and the relative positions between objects scale by the same factor. The cluster is then moved to the center of the artboard.

Bleed is fixed at 3 mm per side, so it targets the artboard dimensions plus 6 mm in both width and height. Because the offset is symmetrical, the center still matches the artboard center.

## Keep aspect and One side only

Switch between these with the radio buttons at the top of the dialog.

- **Keep aspect**: scales while preserving the ratio (default)
- **One side only**: changes just the width or just the height

Selecting "One side only" **dims Ref. side, Area, Artboard, and Bleed**, and clears whichever of them was selected. Those bases only make sense when the aspect ratio is preserved. The three that remain available are Max, Min, and Fixed Size.

Switching modes keeps the currently selected base and recalculates it under the new mode.

## Alignment

Set this with the two panels on the right.

| Align (H) | Behavior |
| --- | --- |
| Left | Matches left edges to the leftmost object |
| Center | Matches center X to the average position |
| Right | Matches right edges to the rightmost object |
| Distribute evenly | Distributes with equal gaps along the vertical axis |
| Zero gap | Stacks vertically with no gaps |

| Align (V) | Behavior |
| --- | --- |
| Top | Matches top edges to the topmost object |
| Middle | Matches center Y to the average position |
| Bottom | Matches bottom edges to the bottommost object |
| Distribute evenly | Distributes with equal gaps along the horizontal axis |
| Zero gap | Butts objects together horizontally |

### Why "Distribute evenly" sits in that panel

The panel name and the distribution direction may look reversed, but this is deliberate.

The "Align (H)" panel holds **the set you need for stacking things vertically**: match the left edges (a horizontal alignment) and space them out vertically (a vertical distribution). It takes both to produce a left-aligned vertical stack, so they live in the same panel.

The "Align (V)" panel is the mirror image: match the top edges and lay them out left to right.

The unit of exclusivity is not the panel but **the coordinate axis being moved**. Everything that moves X (Left / Center / Right, plus the distribution controls in "Align (V)") is mutually exclusive, and everything that moves Y (Top / Middle / Bottom, plus the distribution controls in "Align (H)") is mutually exclusive. **Different axes can be combined**, which is what lets you check both "Left" and "Distribute evenly" inside the "Align (H)" panel.

Note that "Distribute evenly" pins both ends and divides the space between, so it needs **3 or more** objects; "Zero gap" needs **2 or more**.

Alignment survives a change of resize base. After the selection is resized under the new base, the alignment is re-applied to the new sizes.

## The two measurement options

### Measure text by outline bounds

The `geometricBounds` of a text frame is larger than the glyphs themselves, because it includes the leading and descender padding. That gets in the way when you want headline text to match exactly.

Turn this on to measure by **the bounds of the actual outlined glyphs**. It is available (and switched on automatically) only when the selection contains text; otherwise it is dimmed.

### Measure by preview bounds

On by default. When on, measurement uses the visual edges including strokes and effects (`visibleBounds`); when off, it uses the path edges (`geometricBounds`).

**This option affects alignment as well as resizing.** If the two used different bases, an object with a heavy stroke would end up "the same size but not flush at the edges".

## Buttons

- **[Reset]**: reverts size, position, and alignment (the dialog stays open). The resize-base radios and the alignment checkboxes are cleared too
- **[Cancel]**: reverts and closes. **Esc and the window close box do the same**
- **[OK]**: commits the current state and closes

The dialog position is remembered for the current session (it resets when you quit Illustrator).

## Target objects

Any selected PageItem. Text frames and groups get special handling only when measuring outline bounds; text inside groups is collected recursively.

If no document is open, or if nothing is selected, the script shows an alert and exits.

## Notes

"Measure text by outline bounds" works by **duplicating → expanding the appearance → creating outlines → measuring → deleting immediately**, so the original text is never touched. Running that on every measurement would be slow, so results are cached. The cache key includes the object's current bounds (rounded), so resizing changes the key and stale values are never reused.

For the Artboard and Bleed bases, the selection used to be grouped temporarily and scaled as a group. It now scales without any temporary group: the top-left of the cluster is used as the origin, and each item's size and relative position are scaled by the same factor. Because no grouping or ungrouping is involved, there is no structural risk of changing the parent hierarchy or the stacking order.

Objects with zero width or height (stroke-only paths, empty text frames, and so on) cannot be enlarged by `resize()`, and the scale factor would be Infinity, so they are treated as 100% and left as they are.

## Article

[Resizing objects to match — an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n6f35bd4000ec)

## Update history

- v1.4.2 (2026-07-22): Added Japanese and English READMEs, linked from the script's basic info. Fixed: Esc and the window close box now cancel instead of committing the transform. Fixed: Reset now also clears the resize-base radios and the internal base state (clicking an alignment after Reset used to restore the resized geometry). Switching to "One side only" now clears the bases it dims (Ref. side / Area / Artboard / Bleed). Distribution (evenly / zero gap) now moves by delta so it follows "Measure by preview bounds". "Distribute evenly" now requires 3+ objects and "Zero gap" 2+. Added a guard for when no document is open. One-side-only resize now anchors at the top-left like every other mode
- v1.4.1 (2026-07-04): No base is selected when the dialog opens, so it performs no transform on show; resolved the duplicate "Base" label (panel = "Resize base" / row = "Ref. side"); added tooltips to non-obvious options. Fixed a bug where clicking an alignment checkbox wiped the other axis's alignment; alignment now follows "Measure by preview bounds" so stroked and effected objects align by their visual edges
- v1.4.0 (2026-07-04): Overall refactor (shared UI layout, localization, renaming, dead-code removal, split resize and alignment functions). The aspect / one-side toggle keeps the current base; Reset also clears alignment; guarded against Infinity on zero width or height. Artboard and Bleed resizing no longer uses a temporary group, protecting the parent hierarchy and stacking order; added Right and Bottom alignment; reorganized alignment into two panels
- v1.0.0 (2025-04-05): Initial version
