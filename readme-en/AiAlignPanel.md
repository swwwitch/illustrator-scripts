# AiAlignPanel

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAlignPanel.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/AiAlignPanel.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAlignPanel.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that aligns the selection to the artboard.

- Align: seven icons run immediately on click — horizontal (left / center / right), both axes at once, and vertical (top / center / bottom). A multiple selection is grouped temporarily so it moves as one
- Margin: a uniform inset from all four artboard edges (unit follows the ruler). Only the left/right/top/bottom alignments are affected; the three centered ones do not move
	- Option-click ignores the margin and sits flush against the artboard edge
	- The field steps with ↑↓ (Shift = ±10, snapping to multiples of 10 / Option = ±0.1); negative values are clamped to 0
- Options
	- Preview Bounds: align using the visible bounds, including stroke and effects (off by default)
	- Align to Glyph Bounds: toggles point type and area type together (on by default); Option-click treats it as on even when unchecked
	- Change Justification: when centering horizontally, also center the justification of a lone single-line text object (on by default); dimmed while the selection is not text
- Keyboard: Esc closes the palette
- A status line at the bottom reports the result (aligned / nothing selected / error)

DOM work (selection, alignment, the margin offset, and writing the preferences) is delegated to the main engine via BridgeTalk.

### How it works

1. Check the document and the selection
2. If characters are selected, reselect the text objects instead
3. Abort if the selection spans multiple layers
4. Write Preview Bounds and Align to Glyph Bounds to the preferences, as checked
5. Make the artboard holding the selection the active one
6. When centering horizontally, center the justification of a lone single-line text object
7. Group a multiple selection, run the align command, move inward by the margin, then ungroup

### Notes

- The align target (selection / key object / artboard) follows the Align panel. **With "Align to Selection" active nothing moves, because grouping reduces the target to a single object.** Set "Align to Artboard" before running.
- The margin is applied by moving inward after the align command has snapped the object to the artboard edge. The offset is a fixed amount, so it stays exact with either preview or geometric bounds.
- Preview Bounds and Align to Glyph Bounds are not read back from the preferences at startup. The palette writes its own checkbox states to the preferences on every align, because reading preferences from a persistent palette is not reliable.
- A selection spanning multiple layers is refused because grouping moves every object to the frontmost object's layer, and ungrouping does not restore the original layers.
- Grouping and ungrouping add steps to the undo stack.
- Running it with characters selected in threaded text targets every text object in that story.
- Justification is only centered for single-line text. Changing it on multi-line text would shift every line and alter the appearance significantly; area type wrapped onto two or more lines is excluded as well.
- The Change Justification checkbox is dimmed based on whether the selection is text. Illustrator has no timer API, so the selection is re-read when the palette regains focus and when the mouse moves over it.
- In a document with multiple artboards, the one overlapping the selection most becomes active. If the selection overlaps none, the artboard nearest to its center is used.
- Re-running the script closes the open palette and rebuilds it (this also prevents a second instance), so edited code can be applied by simply running it again.
- For a one-shot centering command see [CenterAlignAsGroup](CenterAlignAsGroup.md), or [VerticalCenterAlignAsGroup](VerticalCenterAlignAsGroup.md) for vertical centering only.

### Script info

- Version: v1.0.0

### Changelog

- v1.0.0 (20260823) : Initial release
