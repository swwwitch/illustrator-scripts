# SwapNearestItem.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-SwapNearestItem.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/SwapNearestItem.jsx)


[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that swaps the position of a selected object with the nearest object in a specified direction (right, left, up, or down), maintaining a natural visual layout.
- Supports special handling for multi-selection and considers widths, heights, and gaps.

### Main Features

- Swap based on shortest distance in four directions
- Adjust positions naturally considering widths, heights, and gaps
- Special handling for manual swaps with two selections
- Japanese and English UI support

### Process Flow

1. Check document and selected objects
2. Find the nearest object in the specified direction
3. Swap positions considering sizes and gaps
4. Handle multi-selection with center or edge-based swaps

### Update History

- v1.0.0 (20250610): Initial release
- v1.0.1 (20250612): Added exclusion of group/compound path children and lock support
- v1.0.2 (20250613): Refactored with getCenter() and getSize()
- v1.0.3 (20250614): Added temporary group handling for multiple selection
- v1.0.4 (20250615): Removed temporary grouping, cleaned up logic