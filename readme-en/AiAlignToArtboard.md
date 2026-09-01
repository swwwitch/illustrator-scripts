# Align to guides or artboard edges

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAlignToArtboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/AiAlignToArtboard.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAlignToArtboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that aligns the selection to the artboard. A 3×3 grid of buttons moves the selection in eight directions, stepping the destination outwards on each press: the guide, the artboard edge, then the bleed. It also draws margin and division guides. When the Align panel is not set to Align to Artboard, that is detected and reported, and nothing is moved.

- Buttons: a 3×3 grid (eight around the edge plus the middle) with two centre buttons below it; every one runs immediately on click
	- The eight outer buttons step the destination outwards on each press: the guide they would run into in that direction, then the artboard edge, then the bleed. A diagonal moves to both edges at once. A multiple selection is grouped temporarily so it moves as one, and Change Justification applies
	- The middle button centres on both axes. **Option-click centres inside the margin** instead of on the artboard
	- The two buttons below centre on one axis each; Option-click centres that axis inside the margin
	- Each icon shows a rule at the destination edge with the object pushed against it; on the three centred ones the rule runs through the object instead
- Margin Guide: the inset from each artboard edge, set individually for top, bottom, left and right (unit follows the ruler). Only the left/right/top/bottom alignments are affected; the three centered ones and the arrow buttons do not use it
	- Option-click ignores the margin and sits flush against the artboard edge
	- Pressing again once the selection already sits at the margin carries it on to the artboard edge beyond it (and it stays there on any further press)
	- The fields step with ↑↓ (Shift = ±10, snapping to multiples of 10 / Option = ±0.1); negative values are clamped to 0
	- The four fields sit in a cross (3×3) with Link in the middle. While it is on, Top alone decides all four sides and the other three are dimmed (on by default); turn it off to set them individually
	- While Add Guides is off, all four fields and Link are dimmed together (their values are kept)
	- The starting value depends on the ruler unit (5 for mm, 20 for pt / px / Q, 0.5 for cm, 0.25 for inches). Until a value is typed in, changing the ruler unit swaps in that unit's starting value
	- Turning Add Guides on turns Link on as well and levels the four values off, since an even inset is the usual case
	- The ruler unit is shown once in the panel title ("Margin Guide (mm)") rather than beside each field
	- Add Guides: draws a rectangle guide inset from the artboard edges by the margin. That guide is a destination for the arrow buttons as well
	- Keep Guides: leaves the guide in place when the palette closes (off by default). Dimmed while no guide is drawn at all, but its own state is preserved. It covers the division guides and the artboard edges too
- Division Guides: draw guides that split the area inside the margin into rows and columns. Choose None, Cross or Custom
	- None: no division guides (default)
	- Cross: one horizontal and one vertical guide through the centre of the area inside the margin
	- Custom: split into the given number of Rows and Columns; a count of 1 draws no guide on that axis
	- Gutter: the space between two rows or two columns. Above zero, each split gets a guide on both sides (unit follows the ruler; available under Custom only)
	- Extension: how far the division guides reach outside the artboard. At zero they span the artboard's width and height
	- Artboard Edges: also draw guides on the four edges of the artboard, reaching outside by the extension. It works independently of the division settings
	- The guides are drawn on the "_guide" layer under the name `AiAlignToArtboard-divide`. As with the margin guide, every redraw deletes the existing guides of that name first. They are destinations for the arrow buttons as well
- Measurement Options
	- Preview Bounds: align using the visible bounds, including stroke and effects (off by default)
	- Align to Glyph Bounds: toggles point type and area type together (on by default); Option-click treats it as on even when unchecked. Symbol instances are covered too — the type inside them is measured and aligned to its glyph bounds
- Destination
	- Align to Bleed: adds the bleed outside the artboard as the **stop after the artboard edge** (off by default; the field beside it starts at 3 mm). Being a print value it never follows the ruler unit and is always read as millimetres
	- Align per Artboard: aligns each selected object to **the artboard it sits on** (off by default). Objects sharing an artboard still move together as one block, and the arrow buttons follow the same rule. It is dimmed in a document with only one artboard
	- Both are written to the preferences only for the duration of an align, and the previous values are restored afterwards
	- Change Justification: match the justification of a lone single-line text object to the horizontal alignment (on by default) — left align to left, horizontal center to centered, right align to right. Vertical alignments leave it alone. Dimmed while no text is selected
- Keyboard: Esc closes the palette; ↑↓←→ fire the four side move buttons (the diagonals have no key); C centres horizontally, M vertically and X on both axes; B toggles Align to Bleed
- A status line at the bottom reports the result (aligned / moved / nothing selected / wrong align target / error). Text too long for the fixed width can be read in full by hovering over it

DOM work (selection, alignment, moving, the margin offset, and reading/writing the preferences) is delegated to the main engine via BridgeTalk.

### How it works

The centred buttons:

1. Check the document and the selection
2. If characters are selected, reselect the text objects instead
3. Abort if the selection spans multiple layers
4. Make the artboard holding the selection the active one
5. Take the stop after the one the selection currently sits on (margin → artboard edge → bleed) as this run's destination
6. Remember the current preferences, then write Preview Bounds and Align to Glyph Bounds as checked and redraw so the align commands pick them up
7. On a horizontal alignment, match the justification of a lone single-line text object
8. Group a multiple selection, run the align command, confirm the align target (restore the justification and stop if it is not the artboard), move inward by the margin, then ungroup
9. Redraw, then restore the remembered preferences
10. If Align to Glyph Bounds is on and the selection contains a text object or a symbol, measure the glyphs and cancel out any offset from the target
11. If Add Guides is on, redraw the rectangle guide at the margin

Arrow buttons:

1. Check the document and the selection (if characters are selected, reselect the text objects instead)
2. Make the artboard holding the selection the active one
3. Measure the selection bounds as Preview Bounds and Align to Glyph Bounds are checked (without writing the preferences)
4. Look for a guide that edge would run into in that direction; use it if found, otherwise the artboard edge (a diagonal looks on both axes)
5. Move the selected objects as they are, by the distance that puts that edge on the destination (a diagonal applies both axes in one move)

### Notes

- The align target (selection / key object / artboard) follows the Align panel. **With "Align to Selection" active nothing moves, because grouping reduces the target to a single object.** Set "Align to Artboard" before running.
- Any target other than the artboard is detected and reported as "Set Align To: Artboard." — **neither the margin offset nor the glyph-bounds correction is applied**, so nothing moves and any justification change is rolled back. The check runs only when an align moved nothing: the selection is nudged 4 pt and aligned again to see whether it comes back. An object already in the right place ends up exactly where it started.
- The margin is applied by moving inward after the align command has snapped the object to the artboard edge. The offset is a fixed amount, so it stays exact with either preview or geometric bounds.
- Which of the four margins applies follows the edge being aligned to: Left for a left align, Right for a right align, Top for a top align and Bottom for a bottom align. None of the three centered alignments use one.
- The four edge alignments and the arrow buttons **step outwards on each press**: a guide (the margin guide included), then the artboard edge, then the bleed. Stops with no value are skipped, and once out at the last one further presses leave the selection where it is. The check measures bounds the same way Preview Bounds and Align to Glyph Bounds do. The three centered alignments never use the margin and are unchanged, and Option-click skips the steps and goes straight to the artboard edge.
- Align to Bleed draws no guide; it only adds one more stop to that sequence.
- With Align per Artboard on, the selection is **bucketed by artboard** and each bucket is aligned or moved on its own. Within a bucket the objects are still grouped temporarily and move as one block, and the original selection is restored at the end. If one bucket fails, the run stops there and reports why.
- **The bleed value cannot be read from Illustrator.** The document's bleed setting is not exposed in the DOM — it appears on neither `Document` nor `Artboard`, and is not recorded in the XMP — so it has to be typed in.
- Preview Bounds and Align to Glyph Bounds are not read back from the preferences at startup, because reading preferences from a persistent palette is not reliable. The checkbox states are applied only for the duration of an align and **the previous preferences are restored afterwards**. Preview Bounds in particular is an application-wide setting, so running an align never leaves it changed.
- `app.redraw()` is called right after the preferences are written and right before they are restored, so the align command does not run against the old values.
- When Align to Glyph Bounds is on and the selection contains a text object or a symbol, the result is not left to the align command: **after aligning, the glyph bounds are measured and any offset from the target is cancelled out**. Each text object is duplicated, converted with `createOutline()` and measured (the duplicate is then deleted) while everything else is measured with its ordinary bounds; the selection is then moved by the difference between the bounds enclosing it all and the artboard edge plus the margin. When the align command already honoured the setting the difference is 0, so nothing moves twice.
	- Thanks to this correction, **glyph bounds also work for a multiple selection**, which the temporary grouping would otherwise defeat.
	- **Symbol instances are measured this way too.** Illustrator's Align to Glyph Bounds has no effect on symbols, so the symbol is duplicated onto the layer, unlinked, and the type inside it is converted with `createOutline()` before the bounds are measured (the duplicate is then deleted and the selection restored). A symbol with no type inside falls back to its ordinary bounds.
	- Only top-level text objects and symbol instances are measured this way; text nested inside a group falls back to ordinary bounds.
	- Preview Bounds on measures the outline's `visibleBounds`, off measures its `geometricBounds`.
- A selection spanning multiple layers is refused because grouping moves every object to the frontmost object's layer, and ungrouping does not restore the original layers.
- Grouping and ungrouping add steps to the undo stack. They also collapse the stacking order of a multiple selection into a contiguous run at the frontmost object's position.
- Running it with characters selected in threaded text targets every text object in that story.
- Justification is only changed for single-line text. Changing it on multi-line text would shift every line and alter the appearance significantly; area type wrapped onto two or more lines is excluded as well.
- The Change Justification checkbox is dimmed based on whether the selection is text, and Align per Artboard based on whether the document has more than one artboard. Illustrator has no timer API, so the selection is re-read when the palette regains focus and when the mouse moves over it. If the selection has changed by then, the previous result is cleared from the status line.
- In a document with multiple artboards, the one overlapping the selection most becomes active. If the selection overlaps none, the artboard nearest to its center is used.
- The ↑↓←→ keys do exactly what pressing the matching move button does, and C / M / X do the same for the centring icons. B toggles the Align to Bleed checkbox. **A focused field takes the keystroke first**, so typing a margin or a bleed never moves the selection by accident, and nothing is stolen while Command (Ctrl) is held.
- The outer buttons move the selection so that its edge on that side lands on a guide; with no guide it goes to the artboard edge. The margin fields themselves are not used. With Add Guides on, the first press stops at the margin guide, a second carries on to the artboard edge, and with Align to Bleed on a third reaches the bleed beyond it.
- The four diagonals work out both the horizontal and the vertical distance **from the bounds before anything moves**, then apply them in a single move. Moving one axis first would run the other axis's guide search from the new position.
- Only guides **inside the active artboard** are considered. Besides ruler guides, a rectangle (or any other shaped) guide contributes its two opposing edges — left and right for a horizontal move, top and bottom for a vertical one.
- **A guide that does not overlap the selection on the perpendicular axis is ignored**, because moving would never run into it. Moving left or right only considers vertical guides whose vertical extent overlaps the selection, and vice versa. Ruler guides extend beyond the artboard, so they always overlap. Edges that merely touch (zero overlap) do not count.
- A guide the edge already sits on is not used as a destination (anything within 0.001 pt counts as the same position). A hand-snapped edge and its guide can differ by about 1e-12, which would otherwise make the object pick its own position and never move.
- With no guide in the direction of travel the selection goes to the artboard edge — **except when it spans the guides entirely (the object is larger than the gap between them)**, in which case it snaps back to the nearest guide behind. This is what lets an object larger than the margin guide line up with the margin.
- A clipping group is measured by **the bounds of its mask path**. The bounds Illustrator's DOM reports for a clipping group cover the whole content and ignore the mask, which would otherwise send the arrow buttons to the wrong place (the three centred buttons already line up with the mask, because the align command does). A compound path mask is handled as well; with no mask found the ordinary bounds are used.
- The outer buttons measure bounds the same way as the centred ones (Preview Bounds / Align to Glyph Bounds; type is duplicated and outlined to measure it). They never write the preferences, but a multiple selection is grouped temporarily just as it is when aligning.
- The guide is drawn on the "_guide" layer (created if missing; unlocked and shown) as a rectangle named `AiAlignToArtboard-margin`. Every redraw deletes the existing guide of that name first, so there is **always exactly one, on the active artboard**. Aligning on a different artboard moves the guide there.
- The guides are redrawn when Add Guides is toggled, when a margin or division field commits (Enter or focus loss), when the division mode or Artboard Edges is switched, and when an align runs. While stepping with ↑↓ only the dimming follows along, so Illustrator is not queried on every keypress.
- Unchecking Add Guides deletes the guide. **Closing the palette deletes it too** (whether closed with Esc or rebuilt by re-running the script); turn on Keep Guides to leave it behind. A guide moved off the "_guide" layer is left alone.
- The selection kind, the ruler unit and the selection count come back in a single round trip. When Illustrator does not answer, or answers with an error, the dimming and the unit label are left as they are rather than rewritten from an unreliable value.
- DOM work is delegated to the main engine over BridgeTalk; where BridgeTalk is unavailable it runs in the palette's own engine instead.
- Re-running the script closes the open palette and rebuilds it (this also prevents a second instance), so edited code can be applied by simply running it again.
- For a one-shot centering command see [CenterAlignAsGroup](CenterAlignAsGroup.md), or [VerticalCenterAlignAsGroup](VerticalCenterAlignAsGroup.md) for vertical centering only.

### Script info

- Version: v1.2.0

### Changelog

- v1.2.0 (20260901)
	- Add a Division Guides panel: None / Cross / Custom, with rows, columns, gutters and an extension
	- Add Artboard Edges, drawing guides on the four edges of the artboard
	- Fold every button into the 3×3 grid: drop the row of seven align icons, put Align Center on Both Axes in the middle and Horizontal / Vertical Center below it
	- Give the eight outer buttons the bleed stop, Change Justification, and the temporary grouping of a multiple selection
	- Option-click on the three centred buttons now centres inside the margin
	- Redraw the icons: no more arrows — every icon is a rule at the destination edge with the object pushed against it
	- Dim the four margin fields and Link while Add Guides is off
	- Fix a path that failed to become a guide being left behind as a filled object
	- Fix the stored settings being overwritten with defaults by controls that did not exist yet while the palette was being built
- v1.1.0 (20260901)
	- Add four diagonal arrow buttons (top left, top right, bottom left, bottom right) for eight directions in a 3×3 grid
	- Step the destination outwards on each press: margin, then the artboard edge, then the bleed
	- Split the margin into top/bottom/left/right fields (laid out in a cross, Link in the middle) and use the margin for the edge being aligned to
	- Add an Align to Bleed option (mm field, 3 by default)
	- Add an Align per Artboard option
	- Add keyboard shortcuts: ↑↓←→ for the move buttons, C / M / X for the centring icons, B to toggle Align to Bleed
	- Rework the panels into Measurement Options (how bounds are measured) and Align Options (bleed and per-artboard), with Margin Guide moved to the full width of the bottom row and its unit in the panel title
	- Fix Align to Glyph Bounds having no effect on symbol instances (duplicate, break the link, and measure the type inside as outlines)
	- Fix the arrow buttons moving a clipping group by its unclipped bounds instead of its mask path
- v1.0.3 (20260901) : Add four arrow buttons that move the selection to a guide or the artboard edge; lay the palette out in two columns (arrow buttons on the left, Margin Guide and Options on the right); rename the panel to Margin Guide and the checkbox to Add Guides, and retitle the palette "Align to Guides or Artboard Edges"
- v1.0.0 (20260824) : Add "Use the margin", Show Guides and Keep Guides (the guide is deleted when the palette closes unless Keep Guides is on); align to glyph bounds by measuring the outlined glyphs, which also makes it work on a multiple selection; detect and refuse an align target other than the artboard; restore the preferences after aligning; fix decimals being impossible to type in the margin field
- v1.0.0 (20260823) : Initial release
