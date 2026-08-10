# AiAdjustVerticalGap

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAdjustVerticalGap.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/AiAdjustVerticalGap.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAdjustVerticalGap.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A docking palette that sets the vertical gap between two selected objects,
with a live preview that updates as you change the settings.

- Targets two selected objects, or a single group containing exactly two objects
- On open, reads the current gap of the two selected objects into the field (nothing moves)
- Keeps the chosen key object (top or bottom) in place and moves the other
- The gap value uses the document's ruler unit (arrow keys: Shift ±10 / Option ±0.1)
- Negative gap values overlap the two objects
- Optional horizontal alignment (none / left / center / right)
- An extra "Offset" value shifts the moving object further horizontally after alignment (positive = right, negative = left; unit follows the ruler), and works even when align is none
- Optional paragraph alignment for text (keep / match align / justify); left alignment works around an Illustrator bug via a temporary resize
- Clip groups measure by their clipping path; preview bounds (stroke/effects) can be toggled
- Record saves the current settings and locks (dims) the panels, switching the button to Edit (click again to unlock)
- While locked, select multiple groups and Apply to batch-apply the recorded settings (each group of two, or two selected objects)
- Closing with an uncommitted preview reverts it
- Keys: T = top / B = bottom, N/L/C/R = none/left/center/right, S/J = match/justify, A = apply, Esc = close (panel shortcuts are disabled while locked)

### Script info

- Version: v1.3.0
