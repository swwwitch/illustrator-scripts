# Quick Preferences

[![Direct](https://img.shields.io/badge/Direct%20Link-AiQuickPrefsPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/AiQuickPrefsPalette.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

You want to turn "Scale Strokes & Effects" on. Or switch to "Preview Bounds" for a moment. Or set the arrow-key increment to 1 mm.

All of these come up over and over while you work, and every time it means opening the Preferences dialog, finding the right category, ticking a box, and clicking OK. Worse, the items are scattered across General, Units, and Selection & Anchor Display — **the handful you actually touch are the ones furthest away**.

So this script collects just those items into a single palette where a click takes effect immediately. It is meant to be left open.

<img alt="" src="" width="50%" />

## How to use

Run the script to open the palette. From there, operating a checkbox or button **applies the change at that instant**.

**There is no OK or Apply button.** The moment you click is the moment it takes effect. The palette also closes with `Esc` (while it is active).

Running the script again while the palette is already open brings the existing palette to the front instead of reopening it.

## Available controls

| Panel | Items | Notes |
| --- | --- | --- |
| Key Input | value + unit | Arrow-key increment (Preferences > General > Keyboard Increment). The unit popup switches the ruler unit |
| Align Options | Preview Bounds | Use bounds including stroke & effects for align/distribute |
| Align to Glyph Bounds | Point Type / Area Type | Align type to glyph bounds |
| Transform Options | Pattern / Corners / Strokes & Effects | Transform patterns; scale corner (live corner) radius and strokes & effects when scaling |
| Transform | Flip Horizontal / Flip Vertical / Rotate 45° | Acts on the selection (see below) |
| Artboard | Show Artboard Name / Show Video Ruler | The video ruler is a toggle button |
| Artboard border | Highlight color / stroke width | Nine colors, widths 1–4 |
| Copy / Paste | Paste without Formatting / Paste Remembers Layers | |
| Drawing | Real-time Drawing & Editing / Refresh Preview | Refresh Preview redraws the GPU preview |

Every control has a tooltip, so hover to see which preference it maps to.

## Option+click to toggle a whole panel

The **Align to Glyph Bounds** and **Transform Options** checkboxes respond to **Option+click by setting every item in that panel to the same state**.

- Align to Glyph Bounds: point type / area type (2 items)
- Transform Options: pattern / corners / strokes & effects (3 items)

A plain click still toggles a single item as before. Transform Options is often turned on or off as a set, so adding Option is all it takes.

## Numeric input for Key Input

The value field responds to the `↑` `↓` keys.

| Key | Step |
| --- | --- |
| `↑` `↓` | ±1 |
| `Shift` + `↑` `↓` | ±10 (snaps to the next multiple of 10) |
| `Option` + `↑` `↓` | ±0.1 |

The value never goes negative — it clamps at 0.

The unit popup does more than change the displayed unit: it **switches the ruler unit (rulerType) itself**. This is the slightly confusing part. Switching units does not change the stored increment in points; only the display is recomputed in the new unit.

The popup lists seven units: in / mm / pt / pica / cm / Q/H / px.

## Transform panel

"Flip Horizontal", "Flip Vertical", and "Rotate 45°" all pivot about **the center of the visible bounds of the whole selection**. Individual objects do not spin around their own centers; the selection flips or rotates as a single block.

The rotation direction is set with the radio buttons on the right (clockwise / counterclockwise). The default is counterclockwise.

Items whose bounds cannot be read or that cannot be transformed (locked, hidden, guides, and so on) are skipped.

## Following external changes

If you change a setting outside the palette — in the Preferences dialog, for example — **clicking the palette (re-activating it) syncs the display**.

The one exception: while the Key Input value field has focus, syncing is skipped so that what you are typing is not overwritten.

## Artboard border color

The dropdown offers nine presets (light blue / salmon pink / green / medium blue / magenta / cyan / light gray / black / yellow).

If the current setting does not match a preset, **the preset with the closest RGB value is shown as selected**. This is not a way to specify an arbitrary color — it is a dropdown for picking a frequently used one quickly.

## Notes

A persistent Illustrator palette loses its DOM connection while it is shown, so all preference **writes** and object transforms are delegated to the main engine via BridgeTalk. **Reads** are safe across engines and are fetched directly and synchronously in the palette.

Flips and rotations combine the three steps "translate → flip/rotate → translate" into a single composite matrix, reducing `transform` to one call per object. The cos / sin values are computed numerically in the palette and embedded into the delegated code.

"Corners" alone is an integer preference (`policyForPreservingCorners`, 1=ON / 2=OFF) rather than a boolean, so the shared group-toggle logic takes a per-item apply function to absorb the difference.

The "Refresh Preview" button toggles the `View using GPU` menu command twice to force a redraw. Changing the artboard border also redraws the canvas via `zoomout` → `zoomin`. Because that toggle snaps the view to the nearest zoom step, the original zoom and center point are saved and restored around it.

## Change log

- v1.8.2 (2026-08-01) Fixed a bug where re-running the script with the palette open silently stopped arrow-key edits to Key Input from being saved. Fixed click-to-sync not working right after launch (while the Key Input field held focus). Display decimals now follow the unit (3 for inches, and so on); the artboard-border redraw restores the original zoom and center point, and the write plus redraw are skipped when nothing changed. Unified the checkbox write path. Cleaned up the header comment (the overview now points to the README) and moved the basic-info block to the standard format.
- v1.8.1 (2026-06-29) Option+click now toggles the whole group for the Align to Glyph Bounds (2 items) and Transform Options (3 items) checkboxes. The linked behavior was unified into linkCheckboxGroup (corners, an integer preference, is absorbed by an apply function).
- v1.8.0 (2026-06-28) Added a "Rotate 45°" button (direction via radios, default counterclockwise, pivoting on the selection center). Flip buttons placed side by side and renamed "Flip Horizontal / Flip Vertical"; panel renamed "Transform". Added an "Align Options" title to the align panel and shortened the artboard panel name. Column spacing adjusted with COLUMN_SPACING.
- v1.7.1 (2026-06-27) Renamed "Other" to "Copy / Paste" and added a "Drawing" panel (Real-time Drawing & Editing moved there). Added a Refresh Preview button to the Drawing panel.
- v1.7.0 (2026-06-27) Added the transform (flip) panel (pivots on the selection center, skips locked/guide items, sped up with a composite matrix). Reworked the layout (left = Key Input / Align / Glyph Bounds, right = Transform Options / Transform). Removed the text/unit panels and canvas color, added "Paste Remembers Layers". Closes with Esc, and skips syncing while editing. Naming cleanup, shared checkbox creation and button-height adjustment.
- v1.6.0 (2026-06-27) Reorganized to the standard format (IIFE, localization structure, block comments). Turned into a palette with BridgeTalk delegation, added the key-input unit popup, click-to-sync, guide/artboard panels, and a two-column layout.
- v1.0 (2025-08-04) Initial version.

### note

- https://note.com/dtp_tranist/n/n41d8dc1961be
