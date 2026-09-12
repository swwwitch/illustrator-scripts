# Reset rotation, shear, scale and aspect ratio

[![Direct](https://img.shields.io/badge/Direct%20Link-ResetTransform.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/ResetTransform.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetTransform.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Safely resets rotation, shear, scale and aspect ratio on placed images, text, rectangles (paths), clip groups and straight paths.
- The bounding box is reset and the item is moved back to its original center, so its apparent position never changes.

### Main Features

- **Placed images / raster**: reset rotation, shear, aspect ratio (matching the smaller axis to the larger one, as a rounded integer percentage), flip (vertical / horizontal) and scale (a given percentage, with a 20% floor), individually or together.
- **Text**: reset rotation, shear, the horizontal / vertical scale (to 100%) and tracking (to 0).
- **Rectangles (four-point paths)**: when the path sits 0.5–44° off an axis, snap it to the nearest axis by the smallest possible rotation.
- **Straight lines (two-point paths)**: the same rule, snapping to the nearest axis (0° or 90°).
- **Clip groups**: reset rotation and flip on both the placed image and the mask path, and apply the uniform-scale delta derived from the placed image to both of them at once.
- **Selection handling**: groups and compound paths are traversed recursively to collect targets. When an object inside a clip group is selected, the topmost clip group is processed once instead.
- **UI**: two columns, one panel per target type, hotkeys (S = scale, F = flip), a numeric scale field (arrow keys ±1, Shift+arrows snapping to multiples of 10), and remembered dialog position and opacity.

### Process Flow

1. Check the document and the selection, collecting targets from inside groups and compound paths, and dim the panels that do not apply
2. Choose the operations in the panels and click Reset
3. Each handler applies its transform, resets the bounding box, and moves the item back to its original center
4. Restore the selection that was active when the script started

### Notes

- Rectangles and straight lines are corrected only when they sit **0.5–44°** off an axis. A deliberate angle such as 45° is left alone.
- A rectangle is returned to the nearest axis by the **smallest** rotation, so its width and height are never swapped.
- Clip groups are judged from their placed image. **A clip group with no placed image (vector artwork only) is not a target**, and if nothing else qualifies you get "No resettable objects are selected."
- Resetting the text ratio and tracking applies to the whole text frame.
- Compound paths are corrected per subpath, so a compound path made of several subpaths can end up distorted.
- The scale value is rounded to an integer percentage and raised to 20% when it falls below that.

### Acknowledgements

Noriaki Fujita

### Update History

- v1.6.1 (20260911): Fixed rectangle correction to use the smallest rotation (width and height are no longer swapped), fixed Shift+Down stalling on multiples of 10, fixed very small scales collapsing to 0%, and clip groups without a placed image are now reported as not resettable; internal cleanup (merged duplicate logic, split the dialog builder, revised naming, added JSDoc)
- v1.6.0 (20260708): Dropped the 0.1 arrow-key step in favour of integers only, lowered the scale floor to 20%, widened the rectangle rotation correction to 44° and added a bounding-box reset, restored the original selection after running, and added a guard for when no document is open
- v1.5.1 (20260708): Internal cleanup (IIFE, a shared localization helper with categorized LABELS, a shared `setupPanel` helper, removal of dead code, clearer variable and function names); behaviour matches v1.5
- v1.5 (20250818): Added flip (vertical / horizontal)

### Article

https://note.com/dtp_tranist/n/n52f6b645bc70

### Script info

- Version: v1.6.1
- Last updated: 2026-09-12
