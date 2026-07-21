# ResetRotation

[![Direct](https://img.shields.io/badge/Direct%20Link-ResetRotation.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/ResetRotation.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that levels the apparent rotation of the selected objects back to horizontal (0°).
- Targets text, placed / embedded images, and rectangles (paths).
- When an image sits inside a clipping group, the parent (host) that actually carries the rotation is detected automatically.
- Vertical text is skipped, and objects that are already nearly horizontal (within the tolerance) are left untouched.

### Main Features

- "Objects to Level" panel: Text / Placed-Embedded Image / Rectangle (Path) checkboxes (all ON by default)
- "Clipping Group" checkbox: when ON, the clip group itself is rotated (scope is fixed to Topmost and cannot be changed from the UI)
- "Keep aspect ratio" in the "Text" panel: resets the character horizontal / vertical scale to 100% (ON by default)
- "Level Tolerance (°)" in the "Correction Options" panel: numeric threshold (clamped to 0.01–10°, default 0.1). Arrow keys ±1, Shift+arrows ±10, Option+arrows ±0.1
- Rotation angle is estimated from the transformation matrix, falling back to path vertices (first segment) for paths
- Mirrored transforms (negative determinant) are taken into account when deciding the rotation direction
- Rotation is applied about the object center (Transformation.CENTER), followed by a per-item "Reset Bounding Box"
- Each host is processed only once even when several selected items resolve to the same host
- Automatic Japanese / English UI

### Workflow

1. Recursively traverse the selection and collect text, images and rectangles by type
2. Estimate the rotation angle from the matrix (or path vertices) and skip anything inside the tolerance as already level
3. Rotate by the required angle about the object center, accounting for mirrored transforms
4. Immediately select the rotated item, run "Reset Bounding Box", then restore the previous selection

### Not Supported

- No open document, or nothing selected
- Vertical text (checked only for text frames outside clipping groups)
- Paths that are not closed 4-point shapes (they are not treated as rectangles)
- Objects already horizontal within the tolerance
- The mask path itself while clipping-group mode is on
- Objects that cannot be rotated because they are locked or hidden (exceptions are suppressed and the item is skipped)

### Update History

- v1.0 (20250815): Initial release
- v1.1 (20250815): Added the tolerance UI and arrow-key increments, clipping-group rotation, mirrored image handling, and per-item Reset Bounding Box after rotation
- v1.2 (20250815): Fixed clip scope to "Topmost" and removed scope selection from the UI
- v1.3 (20250815): Added the text frame aspect-ratio option
