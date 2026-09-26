# Toggle the preferences you use most from a persistent palette

[![Direct](https://img.shields.io/badge/Direct%20Link-AiQuickPrefsPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/AiQuickPrefsPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiQuickPrefsPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

You want to flip "Scale Strokes & Effects". You want Preview Bounds on for a moment. You want the arrow keys to move things by 1mm.

All of these come up constantly, and each time it means opening the Preferences dialog, finding the right category, ticking a box and clicking OK. Worse, the settings are scattered across General, Units and Selection & Anchor Display, so **the handful you actually touch are the ones furthest away**.

This palette collects just those settings onto one panel and applies them the instant you click. It is meant to be left open.

<img alt="" src="" width="50%" />

### Features

| Panel | Item | What it does |
| --- | --- | --- |
| Key Input | Value + unit | Keyboard increment (Preferences > General). The unit popup switches the ruler unit |
| Align Options | Preview Bounds | Use bounds including strokes and effects for align/distribute |
| | Align to Glyph Bounds | Align point type and area type to glyph bounds (both toggled together) |
| Transform Options | Pattern Tiles / Corners / Strokes & Effects | Transform patterns along with the object; scale live-corner radii, strokes and effects when scaling |
| Copy / Paste | Paste without Formatting / Paste Remembers Layers | |
| Drawing | Real-time Drawing & Editing / Refresh Preview | Refresh Preview redraws the GPU preview |
| View | Edges / Artboards / Video Ruler / Canvas Color | All toggle buttons; Edges switches edges and the bounding box together |

Every item has a tooltip, so hovering tells you which preference it maps to.

### Usage

Run the script and the palette opens. From there, every checkbox and button **applies the moment you touch it**.

**There is no OK and no Apply.** The click is the commit. The palette also closes with `Esc` while it is active.

Running the script again while the palette is open brings the existing palette forward instead of opening a second one.

### Options

#### Option-click to toggle a whole group

The three Transform Options checkboxes (Pattern Tiles / Corners / Strokes & Effects) **all take the same state on Option-click**. A plain click toggles just the one you clicked. Since these three usually go on or off together, holding Option is enough.

Align to Glyph Bounds is a single checkbox that always toggles point type and area type together.

#### Typing into Key Input

The value field steps with the `↑` and `↓` keys.

| Key | Step |
| --- | --- |
| `↑` `↓` | ±1 |
| `Shift` + `↑` `↓` | ±10 (snaps to the next multiple of 10) |
| `Option` + `↑` `↓` | ±0.1 |

Values never go negative; they clamp at 0.

The unit popup does more than change how the number is displayed — it **switches the ruler unit (`rulerType`) itself**. This part is easy to miss. Switching units does not change the stored increment in points; only the displayed value is recomputed in the new unit.

The popup lists seven units: in / mm / pt / pica / cm / Q / px.

### Notes

If you change a setting outside the palette, for example in the Preferences dialog, **clicking the palette (re-activating it) syncs the display**. The one exception is while the Key Input field has focus, where syncing is skipped so your in-progress value is not overwritten.

A persistent Illustrator palette loses its DOM connection while it is on screen, so every preference **write** is delegated to the main engine over BridgeTalk. **Reads** are safe across engines, so the palette fetches them directly and synchronously.

Corners is the one item that is an integer preference rather than a boolean (`policyForPreservingCorners`, 1 = on / 2 = off), which the write path handles.

The Refresh Preview button toggles the `View using GPU` menu command twice to force a redraw. The Canvas Color button likewise rewrites `uiCanvasIsWhite` and then repaints the canvas with `zoomout` → `zoomin`.

Flipping and rotating the selection lives in [QuickTransformPalette](QuickTransformPalette.md); artboard names and borders live in [ArtboardDisplayPresetManagerPalette](ArtboardDisplayPresetManagerPalette.md).

### Article (note)

https://note.com/dtp_tranist/n/n41d8dc1961be

### Update History

- v2.2.1 (2026-09-19) Merged AiQuickPrefsPalette-simple.jsx and AiQuickPrefsPalette-SuperSimple.jsx into this script. Flip/rotate moved to QuickTransformPalette.jsx; artboard name and border moved to PresetManagerArtboard.jsx.
- v2.0.4 (2026-07-23) Added the "Open File Handling" button.
- v2.0.3 Narrowed to preference toggling (Key Input / Align Options / Transform Options / Copy & Paste / Drawing). Writes delegated over BridgeTalk, reads fetched synchronously. Added Option-click group toggling and click-to-sync with external changes.
- v1.0 (2025-08-04) Initial version.
