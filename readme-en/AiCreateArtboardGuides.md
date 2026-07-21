# AiCreateArtboardGuides

[![Direct](https://img.shields.io/badge/Direct%20Link-AiCreateArtboardGuides.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/AiCreateArtboardGuides.jsx)

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
- Arrow keys step by ±1, Shift by ±10 (snapping to multiples of 10)
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

- v1.1.0: Current version
