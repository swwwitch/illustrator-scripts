# Quick Preferences (SuperSimple)

[![Direct](https://img.shields.io/badge/Direct%20Link-AiQuickPrefsPalette--SuperSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/AiQuickPrefsPalette-SuperSimple.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- A persistent palette that gathers the Illustrator preferences you toggle most often into a single window, so you no longer have to open the Preferences dialog, hunt for the right category, tick a box, and click OK every time.
- Everything applies **the instant you click** — there is no OK or Apply button.
- This **SuperSimple** edition drops the transform (flip/rotate) and artboard panels of the full version ([AiQuickPrefsPalette](AiQuickPrefsPalette.md)) and focuses on preference toggles. In their place it adds a **Preferences** panel with UI-brightness swatches and an "Open File Handling" button.

### How to use

- Run the script to open the palette, then operate the checkboxes and buttons — each change takes effect immediately.
- Close the palette with `Esc` while it is active.
- Running the script again while the palette is open just brings the existing palette to the front instead of reopening it.

### Panels

| Panel | Items | Notes |
| --- | --- | --- |
| Key Input | value + unit | Keyboard increment (Preferences > General). The unit popup switches the ruler unit |
| Align Options | Preview Bounds | Use bounds including stroke & effects for align/distribute |
| | Align to Glyph Bounds | Align point & area type to glyph bounds (both toggled together) |
| Transform Options | Pattern / Corners / Strokes & Effects | Transform patterns, and scale corner radius / strokes & effects when scaling |
| Copy / Paste | Paste without Formatting / Paste Remembers Layers | |
| Drawing | Real-time Drawing & Editing / Refresh Preview | Refresh Preview redraws the GPU preview |
| Preferences | Interface color (4 steps) | Clicking a swatch opens Preferences (User Interface) |
| | Open File Handling | Button that opens Preferences (File Handling) |

Every control has a tooltip, so hover to see which preference it maps to.

### Option+click to toggle a whole panel

- The **Align to Glyph Bounds** and **Transform Options** checkboxes respond to **Option+click by setting every item in that panel to the same state**.
  - Align to Glyph Bounds: point type / area type (2 items)
  - Transform Options: pattern / corners / strokes & effects (3 items)
- A plain click still toggles a single item as before. Transform Options is often turned on/off as a set, so Option saves the extra clicks.

### Numeric input for Key Input

The value field responds to the arrow keys:

| Key | Step |
| --- | --- |
| `↑` `↓` | ±1 |
| `Shift` + `↑` `↓` | ±10 (snaps to the next multiple of 10) |
| `Option` + `↑` `↓` | ±0.1 |

The value never goes negative — it clamps at 0.

The unit popup does more than change the displayed unit: it **switches the ruler unit (rulerType) itself**. Switching units does not change the stored increment in points; only the displayed value is recomputed in the new unit. The popup lists in / mm / pt / pica / cm / Q/H / px.

### Preferences panel (Brightness / File Handling)

#### Interface color (brightness)

The Preferences panel shows four swatches — Dark / Medium Dark / Medium Light / Light. The step closest to the current value is highlighted with a blue selection border.

Because the interface color cannot be applied directly from a script, **clicking a swatch opens Preferences (User Interface)**. Confirm with the arrow keys + Return in that dialog to apply it.

#### Open File Handling

The "Open File Handling" button at the bottom of the panel opens Preferences (File Handling) directly by running the `FilePref` menu command (Illustrator 2022+; the old combined "File Handling & Clipboard" panel was `FileClipboardPref`).

### Following external changes

If you change a setting outside the palette (e.g. in the Preferences dialog), **clicking the palette (re-activating it) syncs the display**. The one exception: while the Key Input value field has focus, syncing is skipped so it does not overwrite what you are typing.

### Notes

- A persistent Illustrator palette loses its DOM connection while shown, so all preference **writes** are delegated to the main engine via BridgeTalk. **Reads** are safe across engines and are done directly and synchronously in the palette.
- "Corners" is an integer preference (`policyForPreservingCorners`, 1=ON / 2=OFF) rather than a boolean, so the group-toggle logic passes a per-item apply function to absorb the difference.
- The "Refresh Preview" button toggles the `View using GPU` menu command twice to force a redraw.

### Change log

- v2.0.4 (2026-07-23) Current version. Added an "Open File Handling" button (`FilePref`) to the Preferences panel.
- v2.0.3 SuperSimple edition. Focused on preference toggles (Key Input / Align Options / Transform Options / Copy·Paste / Drawing) and added a Brightness panel for the UI interface color. Writes delegated via BridgeTalk, reads fetched synchronously. Supports Option+click group toggling and click-to-sync with external changes.

### note

- https://note.com/dtp_tranist/n/n41d8dc1961be
