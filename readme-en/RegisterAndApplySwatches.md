# RegisterAndApplySwatches

[![Direct](https://img.shields.io/badge/Direct%20Link-RegisterAndApplySwatches.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/RegisterAndApplySwatches.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegisterAndApplySwatches.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script to register fill and stroke colors of selected objects (closed paths, text) as swatches (spot colors) and reapply them immediately.
- Supports RGB and CMYK colors, and reuses existing swatches with the same name.

### Main Features

- Register swatches from RGB or CMYK colors
- Reuse existing swatches with matching names
- Immediately apply spot colors to fills and strokes
- Supports text object character colors
- Recursively process groups and compound paths

### Process Flow

1. Select target objects
2. Detect RGB or CMYK colors
3. Generate swatch names and reuse existing swatches if found
4. Create and apply spot colors to objects

### Update History

- v1.0.0 (20250626): Initial version
