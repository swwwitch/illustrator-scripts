# Set the vertical gap between two objects

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAdjustVerticalGapPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/AiAdjustVerticalGapPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAdjustVerticalGapPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A docking palette that sets the vertical gap between two selected objects,
with a live preview that updates as you change the settings.

- Targets two selected objects, or a single group containing exactly two objects
- On open, reads the current gap of the two selected objects into the field (nothing moves)
- Keeps the chosen key object (top, bottom, or auto) in place and moves the other
- Auto (the default) detects the key object set in Illustrator and uses it as the anchor: only with exactly two selected items; align commands probe temporarily and every item is moved back; falls back to Top when it cannot be detected
- The gap value uses the document's ruler unit; the stepper buttons and arrow keys step it to the next whole number (1.5 → 2; Shift to the next multiple of 10 / Option ±0.1)
- Negative gap values overlap the two objects
- Optional horizontal alignment (none / left / center / right)
- An extra "Offset" value shifts the moving object further horizontally after alignment (positive = right, negative = left; unit follows the ruler), and works even when align is none
- Optional paragraph alignment for text (keep / match align / justify); left alignment works around an Illustrator bug via a temporary resize
- Clip groups measure by their clipping path; preview bounds (stroke/effects) can be toggled
- Record saves the current settings and locks (dims) the panels, switching the button to Edit (click again to unlock)
- While locked, select multiple groups and Apply to batch-apply the recorded settings (each group of two, or two selected objects)
- Closing with an uncommitted preview reverts it
- Keys: T = top / B = bottom / K = auto, N/L/C/R = none/left/center/right, S/J = match/justify, A = apply, Esc = close (panel shortcuts are disabled while locked)

### note

- [【Illustrator】指定した間隔でオブジェクト同士を揃える｜DTP Transit 別館](https://note.com/dtp_tranist/n/n8201294835f9)

### Script info

- Version: v1.4.1

### Update History

- v1.4.3 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.4.2 (2026-10-01) The Record/Apply button row now uses the shared part. Unified the window and panel margins and spacing with the shared layout part
- v1.4.1 (2026-09-29) Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held).
- v1.4.0 (2026-09-27) Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten).
- v1.3.3 (2026-09-26) Renamed the file from `AiAdjustVerticalGap.jsx` to `AiAdjustVerticalGapPalette.jsx`.
- v1.3.2 (2026-09-25) The palette reference is now kept in the persistent engine so CloseAllPalettes.jsx can close it.
