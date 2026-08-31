# Align to guides or artboard edges

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAlignToArtboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/AiAlignToArtboard.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAlignToArtboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that aligns the selection to the artboard. The arrow buttons also move the selection to a guide it would run into, or to the artboard edge. When the Align panel is not set to Align to Artboard, that is detected and reported, and nothing is moved.

- Align: seven icons run immediately on click — horizontal (left / center / right), both axes at once, and vertical (top / center / bottom). A multiple selection is grouped temporarily so it moves as one
- Move: four arrow buttons in a cross move the selection to the guide it would run into in that direction, or to the artboard edge. Unlike aligning, nothing is grouped, so the stacking order and the layers are untouched
- Margin Guide: a uniform inset from all four artboard edges (unit follows the ruler). Only the left/right/top/bottom alignments are affected; the three centered ones and the arrow buttons do not use it
	- Option-click ignores the margin and sits flush against the artboard edge
	- The field steps with ↑↓ (Shift = ±10, snapping to multiples of 10 / Option = ±0.1); negative values are clamped to 0
	- Add Guides: draws a rectangle guide inset from the artboard edges by the margin. That guide is a destination for the arrow buttons as well
	- Keep Guides: leaves the guide in place when the palette closes (off by default). Dimmed while Add Guides is off, but its own state is preserved
- Options
	- Preview Bounds: align using the visible bounds, including stroke and effects (off by default)
	- Align to Glyph Bounds: toggles point type and area type together (on by default); Option-click treats it as on even when unchecked
	- Both are written to the preferences only for the duration of an align, and the previous values are restored afterwards
	- Change Justification: match the justification of a lone single-line text object to the horizontal alignment (on by default) — left align to left, horizontal center to centered, right align to right. Vertical alignments leave it alone. Dimmed while no text is selected
- Keyboard: Esc closes the palette
- A status line at the bottom reports the result (aligned / moved / nothing selected / wrong align target / error). Text too long for the fixed width can be read in full by hovering over it

DOM work (selection, alignment, moving, the margin offset, and reading/writing the preferences) is delegated to the main engine via BridgeTalk.

### How it works

Align icons:

1. Check the document and the selection
2. If characters are selected, reselect the text objects instead
3. Abort if the selection spans multiple layers
4. Make the artboard holding the selection the active one
5. Remember the current preferences, then write Preview Bounds and Align to Glyph Bounds as checked and redraw so the align commands pick them up
6. On a horizontal alignment, match the justification of a lone single-line text object
7. Group a multiple selection, run the align command, confirm the align target (restore the justification and stop if it is not the artboard), move inward by the margin, then ungroup
8. Redraw, then restore the remembered preferences
9. If Align to Glyph Bounds is on and the selection contains a text object, measure the glyphs and cancel out any offset from the target
10. If Add Guides is on, redraw the rectangle guide at the margin

Arrow buttons:

1. Check the document and the selection (if characters are selected, reselect the text objects instead)
2. Make the artboard holding the selection the active one
3. Measure the selection bounds as Preview Bounds and Align to Glyph Bounds are checked (without writing the preferences)
4. Look for a guide that edge would run into in that direction; use it if found, otherwise the artboard edge
5. Move the selected objects as they are, by the distance that puts that edge on the destination

### Notes

- The align target (selection / key object / artboard) follows the Align panel. **With "Align to Selection" active nothing moves, because grouping reduces the target to a single object.** Set "Align to Artboard" before running.
- Any target other than the artboard is detected and reported as "Set Align To: Artboard." — **neither the margin offset nor the glyph-bounds correction is applied**, so nothing moves and any justification change is rolled back. The check runs only when an align moved nothing: the selection is nudged 4 pt and aligned again to see whether it comes back. An object already in the right place ends up exactly where it started.
- The margin is applied by moving inward after the align command has snapped the object to the artboard edge. The offset is a fixed amount, so it stays exact with either preview or geometric bounds.
- Preview Bounds and Align to Glyph Bounds are not read back from the preferences at startup, because reading preferences from a persistent palette is not reliable. The checkbox states are applied only for the duration of an align and **the previous preferences are restored afterwards**. Preview Bounds in particular is an application-wide setting, so running an align never leaves it changed.
- `app.redraw()` is called right after the preferences are written and right before they are restored, so the align command does not run against the old values.
- When Align to Glyph Bounds is on and the selection contains a text object, the result is not left to the align command: **after aligning, the glyph bounds are measured and any offset from the target is cancelled out**. Each text object is duplicated, converted with `createOutline()` and measured (the duplicate is then deleted) while everything else is measured with its ordinary bounds; the selection is then moved by the difference between the bounds enclosing it all and the artboard edge plus the margin. When the align command already honoured the setting the difference is 0, so nothing moves twice.
	- Thanks to this correction, **glyph bounds also work for a multiple selection**, which the temporary grouping would otherwise defeat.
	- Only top-level text objects are measured this way; text nested inside a group falls back to ordinary bounds.
	- Preview Bounds on measures the outline's `visibleBounds`, off measures its `geometricBounds`.
- A selection spanning multiple layers is refused because grouping moves every object to the frontmost object's layer, and ungrouping does not restore the original layers.
- Grouping and ungrouping add steps to the undo stack. They also collapse the stacking order of a multiple selection into a contiguous run at the frontmost object's position.
- Running it with characters selected in threaded text targets every text object in that story.
- Justification is only changed for single-line text. Changing it on multi-line text would shift every line and alter the appearance significantly; area type wrapped onto two or more lines is excluded as well.
- The Change Justification checkbox is dimmed based on whether the selection is text. Illustrator has no timer API, so the selection is re-read when the palette regains focus and when the mouse moves over it. If the selection has changed by then, the previous result is cleared from the status line.
- In a document with multiple artboards, the one overlapping the selection most becomes active. If the selection overlaps none, the artboard nearest to its center is used.
- The arrow buttons move the selection so that its edge on that side lands on a guide; with no guide it goes to the artboard edge. The margin is not used.
- Only guides **inside the active artboard** are considered. Besides ruler guides, a rectangle (or any other shaped) guide contributes its two opposing edges — left and right for a horizontal move, top and bottom for a vertical one.
- **A guide that does not overlap the selection on the perpendicular axis is ignored**, because moving would never run into it. Moving left or right only considers vertical guides whose vertical extent overlaps the selection, and vice versa. Ruler guides extend beyond the artboard, so they always overlap. Edges that merely touch (zero overlap) do not count.
- A guide the edge already sits on is not used as a destination (anything within 0.001 pt counts as the same position). A hand-snapped edge and its guide can differ by about 1e-12, which would otherwise make the object pick its own position and never move.
- With no guide in the direction of travel the selection goes to the artboard edge — **except when it spans the guides entirely (the object is larger than the gap between them)**, in which case it snaps back to the nearest guide behind. This is what lets an object larger than the margin guide line up with the margin.
- Moving measures bounds the same way as aligning does (Preview Bounds / Align to Glyph Bounds; type is duplicated and outlined to measure it), but unlike aligning it never writes the preferences and never groups, so it costs a single undo step.
- The guide is drawn on the "_guide" layer (created if missing; unlocked and shown) as a rectangle named `AiAlignToArtboard-margin`. Every redraw deletes the existing guide of that name first, so there is **always exactly one, on the active artboard**. Aligning on a different artboard moves the guide there.
- The guide is redrawn when Add Guides is toggled, when the margin field commits (Enter or focus loss), and when an align runs. While stepping with ↑↓ only the dimming follows along, so Illustrator is not queried on every keypress.
- Unchecking Add Guides deletes the guide. **Closing the palette deletes it too** (whether closed with Esc or rebuilt by re-running the script); turn on Keep Guides to leave it behind. A guide moved off the "_guide" layer is left alone.
- The selection kind, the ruler unit and the selection count come back in a single round trip. When Illustrator does not answer, or answers with an error, the dimming and the unit label are left as they are rather than rewritten from an unreliable value.
- DOM work is delegated to the main engine over BridgeTalk; where BridgeTalk is unavailable it runs in the palette's own engine instead.
- Re-running the script closes the open palette and rebuilds it (this also prevents a second instance), so edited code can be applied by simply running it again.
- For a one-shot centering command see [CenterAlignAsGroup](CenterAlignAsGroup.md), or [VerticalCenterAlignAsGroup](VerticalCenterAlignAsGroup.md) for vertical centering only.

### Script info

- Version: v1.0.3

### Changelog

- v1.0.3 (20260901) : Add four arrow buttons that move the selection to a guide or the artboard edge; lay the palette out in two columns (arrow buttons on the left, Margin Guide and Options on the right); rename the panel to Margin Guide and the checkbox to Add Guides, and retitle the palette "Align to Guides or Artboard Edges"
- v1.0.0 (20260824) : Add "Use the margin", Show Guides and Keep Guides (the guide is deleted when the palette closes unless Keep Guides is on); align to glyph bounds by measuring the outlined glyphs, which also makes it work on a multiple selection; detect and refuse an align target other than the artboard; restore the preferences after aligning; fix decimals being impossible to type in the margin field
- v1.0.0 (20260823) : Initial release
