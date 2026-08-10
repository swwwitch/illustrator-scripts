# ResetTransform

[![Direct](https://img.shields.io/badge/Direct%20Link-ResetTransform.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/ResetTransform.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetTransform.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Safely resets rotation, shear, scale and aspect ratio on placed images, text, rectangles (paths), clip groups and straight paths.
- A bounding-box reset plus repositioning from the top-left keeps the apparent position stable.

### Main Features

- **Placed images / raster**: reset rotation, shear, aspect ratio (matching the smaller axis to the larger one, as a rounded integer percentage), flip (vertical / horizontal) and scale (a given percentage, with a 20% floor), individually or together.
- **Text**: reset rotation, shear and the horizontal / vertical scale (to 100%).
- **Rectangles (four-point paths)**: snap a near angle (within ±44°) so the rotation becomes 0° or 90°.
- **Straight lines (two-point paths)**: snap to the nearest axis (0° or 90°).
- **Clip groups**: reset rotation and flip on the children (placed/raster items **and the mask path**), and apply the uniform-scale delta derived from the placed image to both the image and the mask path.
- **UI**: two columns, one panel per target type, hotkeys (S = scale, F = flip), a numeric scale field (arrow keys ±1, Shift+arrows in steps of 10), and remembered dialog position and opacity.

### Process Flow

1. Check the document and the selection, analyse what is selected, and dim the panels that do not apply
2. Choose the operations in the panels and click Reset
3. Each handler applies its transform, resets the bounding box, and repositions from the top-left
4. Restore the selection that was active when the script started

### Acknowledgements

Noriaki Fujita

### Update History

- v1.6.0 (20260708): Dropped the 0.1 arrow-key step in favour of integers only, lowered the scale floor to 20%, widened the rectangle rotation correction to 44° and added a bounding-box reset, restored the original selection after running, and added a guard for when no document is open
- v1.5.1 (20260708): Internal cleanup (IIFE, a shared localization helper with categorized LABELS, a shared `setupPanel` helper, removal of dead code, clearer variable and function names); behaviour matches v1.5
- v1.5 (20250818): Added flip (vertical / horizontal)

### Article

https://note.com/dtp_tranist/n/n52f6b645bc70

### Script info

- Version: v1.6.0
- Last updated: 2026-07-08
