# Fit artboard size to your objects

[![Direct](https://img.shields.io/badge/Direct%20Link-FitArtboardWithMargin.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/FitArtboardWithMargin.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitArtboardWithMargin.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Adjusts artboard size by operation (fit to objects / expand the artboard), target (current / all artboards) and size (width & height). The artboard preview updates as you change margins and rounding, and OK commits exactly what you see.

<img alt="The Adjust Artboard Size dialog" src="../png/ss-574-1008-144-20260910-173627.png" width="50%" />

### Main Features

- Two operations: fit to the selected objects' bounds, or grow/shrink the artboards themselves
- Target the current artboard or all artboards
- Adjust width only or height only (the axis you switch off keeps its original size)
- Run Fit with nothing selected to resize every artboard to the objects it contains
- Separate vertical and horizontal margins, with a link icon that mirrors the vertical value
- Margin unit follows the ruler unit, with per-unit defaults (mm=5, px=20, pt=10)
- Switch between preview bounds (including strokes and effects) and geometric bounds
- Rounding for artboard position and size: optimize to pixel grid / round in the current unit / do nothing
- Live preview; Cancel restores the state from when the dialog opened
- The stepper buttons left of each field and the arrow keys change the value: to the next whole number (1.5 → 2), Shift snaps to multiples of 10, Option steps by 0.1
- Remembers the last settings and dialog position for the session
- Japanese / English UI

### Usage

1. Pick an artboard, or select the objects you want to fit to.
2. Run the script.
3. Choose the operation, target and size under Adjustment basis, then enter the margins.
4. Check the preview and click OK.

With Expand artboard, the margin is added to the artboard's own size; enter a negative value to shrink it.

Width and Height can be switched on independently. Option-click one of them to turn that one on and the other off.

### Options

**Adjustment basis**

| Item | Description |
| --- | --- |
| Operation | Fit to objects / Expand artboard |
| Target | Current artboard / All artboards |
| Size | Which dimensions to change (width & height). The one you switch off keeps its original size |

**Margin**

| Item | Description |
| --- | --- |
| Vertical | Amount applied to the top and bottom (editable while Height is on) |
| Horizontal | Amount applied to the left and right (editable while Width is on) |
| Link (icon) | Applies the vertical value to horizontal as well |
| Preview bounds | On measures the visible bounds incl. strokes and effects; off measures geometric path bounds (Fit only) |

**Artboard size fine-tuning**

| Item | Description |
| --- | --- |
| Optimize to pixel grid | Rounds position and size to integer pixels |
| Round values in current unit | Rounds position and size to integers in the current ruler unit |
| Do nothing | Applies the measured values without rounding |

Rounding rounds X, Y, width and height once each, then rebuilds the right/bottom edges as X+width and Y−height, so no value is rounded twice.

### Notes

- Fit + Current artboard requires a selection. With nothing selected, the target is locked to All artboards.
- Fit + All artboards ignores the selection and uses the objects overlapping each artboard. Locked, hidden and guide objects — and objects on locked or hidden layers — are excluded, and artboards with no objects are left untouched.
- Text is measured by outlining a duplicate, so the original text frames are never touched (ID, stacking order, name and tags are preserved).
- Clip groups are measured by their clipping path.
- Any artboard whose width or height would become zero or less is skipped, and a message is shown.
- Settings are kept for the session only and reset when Illustrator restarts.

### Credits

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### Article

https://note.com/dtp_transit/n/n15d3c6c5a1e5

### Changelog

- v1.0 (2025-04-20): Initial version
- v1.1 (2025-07-08): UI improvements, updated default point value
- v1.2 (2025-07-09): UI improvements and bug fixes
- v1.3 (2025-07-10): UI improvements, added panel and radio buttons
- v1.9.2 (2026-09-10): Added width/height checkboxes to Adjustment basis; merged FitArtboardHeight.jsx so that running Fit with nothing selected resizes every artboard to the objects it contains. Also fixed preview restore on a partial failure, group-level effects being dropped from measurements, outlining failures aborting the run, and the first-run dialog appearing off-screen
- v1.10.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.10.1 (2026-09-28): Replaced the Link checkbox with a link icon
- v1.10.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.10.2 (2026-09-28): The button row is now built with the shared part
- v1.10.2 (2026-09-28): Clip groups are now measured by their mask (text and compound-path masks included; overlap with the artboard is also judged by the mask)
- v1.10.3 (2026-09-29): Dialog opacity changed to 98%
- v1.10.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.10.5 (2026-10-01): The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.10.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
