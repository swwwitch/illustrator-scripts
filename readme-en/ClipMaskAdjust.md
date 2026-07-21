# ClipMaskAdjust

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipMaskAdjust.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipMaskAdjust.jsx)

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
- Numeric fields respond to arrow keys: ±1, Shift+arrows ±10, Option+arrows ±0.1
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

- ClipMaskAdjust-v3 (Auto-Preview): updated 2026-01-03
