# Create guides from artboards and convert ruler guides

[![Direct](https://img.shields.io/badge/Direct%20Link-AiCreateArtboardGuides.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/AiCreateArtboardGuides.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCreateArtboardGuides.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that organizes and creates guides relative to artboards.
- Three groups (ruler-guide conversion, center guides, edge guides) are configured together in a single dialog.
- All created guides are collected on a "_guide" layer (created if missing; unlocked and made visible when reused).
- Live preview follows every setting change.

### Main Features

- **Convert ruler guides** (the panel title shows the number of convertible guides)
  - Detects straight ruler guides overlapping an artboard and redraws them as artboard-based straight guides
  - "Extend" sets how far the guides reach beyond the artboard edge (0 = flush with the edge)
  - "All artboards" on: target every artboard the guide overlaps; off: only the first one
  - Turning the master checkbox off skips conversion; with zero targets the section is disabled automatically and a note is shown
- **Center & Edge Guides**
  - "Draw vertical center guide" / "Draw horizontal center guide" add guides at the artboard center (off by default)
  - Turning the edge master on creates guides on the top / left / right / bottom edges (master off by default, the four per-edge checkboxes on by default)
  - "Extend Beyond Edge" sets how far the edge guides extend past the artboard corners (default equivalent to 10 mm)
  - "All artboards" on: draw on every artboard; off: the active artboard only
- **Preview** (on by default): colored lines are drawn on a dedicated layer and replaced by real guides on commit; originals targeted for conversion are hidden temporarily
- Entered values are treated in the current ruler unit (rulerType) and converted to points
- The number fields step with the ∧∨ buttons or the arrow keys to the next whole number (1.5 → 2); Shift snaps to the next multiple of 10, Option steps by 0.1
- Automatic Japanese / English UI, with tooltips on every option

### Workflow

1. Collect every guide in the document and pre-detect the straight guides overlapping an artboard as conversion targets
2. Configure conversion, center, and edge settings in the dialog (the preview re-renders on every change)
3. On OK, remove the original guides being converted and create artboard-based guides in their place
4. Then create edge and center guides on the target artboards (all on the "_guide" layer)

### Not Supported

- No open document (an alert is shown and the script exits)
- Hidden guides
- Guides that are neither vertical nor horizontal (diagonal)
- Guides that do not overlap any artboard
- Documents with no guides at all can still run the center / edge creation

### note

https://note.com/dtp_tranist/n/n56d9c936a364

### Update History

- v1.2.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.2.5 (2026-10-01): The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.2.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.2.3 (20260929): Dialog opacity changed to 98%
- v1.2.2 (20260928): The unit label for the ha ruler unit is now "H" instead of "Q/H". The button row is now built with the shared part
- v1.2.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.1.0: Current version
