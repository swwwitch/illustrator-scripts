# Adjust the mask and contents of a clip group

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipMaskAdjust.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipMaskAdjust.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipMaskAdjust.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that adjusts both the mask path and the contents of a clip group (clipping mask).
- Every dialog change is applied to the canvas immediately as an auto-preview.
- Numeric fields follow the current ruler unit (rulerType).
- For safety, Undo / history handling is not implemented — Cancel does not roll back what the preview already applied.

### Main Features

- "Anchor": a 3x3 grid of radio buttons that sets where the contents are aligned
- "Nudge": X / Y offsets in the current unit, applied on top of the anchor position
- "Fit & Scale"
  - Proportions (Fill): scale the contents to cover the mask
  - Proportions (Fit): scale the contents to fit inside the mask
  - Keep Size: align only, leaving the content size untouched
  - Set Scale: type a percentage directly (pre-filled with the current scale)
- "Mask Path": None / Fit to Content / Square (built from the shorter side, centered)
- "Round Corners": applies the Round Corners effect to the whole clip group. The default radius is (mask width + mask height) / 25, rounded up in the current unit
- "Circle": enabled only while "Square" is selected. Turning it on enables Round Corners and sets the radius to half the shorter side
- Changing the radius clears the existing effect before reapplying it, so effects never stack
- Anchor points can be switched from the keyboard (q/w/e, a/s/d, z/x/c)
- Numeric fields step with the ∧∨ buttons or the arrow keys to the next whole number (1.5 → 2); Shift snaps to the next multiple of 10, Option steps by ±0.1
- Automatic Japanese / English UI

### Workflow

1. Split the selected clip group into its mask path and its contents
2. Reshape the mask path according to the chosen mode (Fit to Content / Square)
3. Apply the Round Corners effect to the clip group when it is enabled
4. Scale the contents against the mask path's visibleBounds and position them using the anchor and nudge values

### Not Supported

- Nothing selected (an alert is shown and the script exits)
- Objects that are not clip groups (a GroupItem with clipped = true)
- Clip groups with no mask path or with no contents
- Undo / history restore (preview results remain even after Cancel)

### Update History

- v3.1.3 (20260929): Dialog opacity changed to 98%
- v3.1.2 (20260928): Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held)
- v3.1.2 (20260928): The nine reference point radio buttons were replaced with the shared 3×3 picker
- v3.1.2 (20260928): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure
- v3.1.2 (20260928): The button row is now built with the shared part
- v3.1.2 (20260928): Clip groups are now measured by their mask (compound-path and text masks included)
- v3.1.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v3.1.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v3.0.2 (20260927): Added tooltips to the English UI. Unit conversion now covers every ruler unit. Values are no longer re-rounded while typing, and the script works when the first selected item is not a clip group
- ClipMaskAdjust-v3 (Auto-Preview): updated 2026-01-03
