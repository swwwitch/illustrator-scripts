# ApplySwatchesToSelection

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplySwatchesToSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ApplySwatchesToSelection.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script that applies colors from selected swatches or all process swatches to selected objects or text.
- You can choose to apply colors sequentially or randomly, and assign colors per character or per object.

### Main Features

- Uses selected swatches from the Swatches panel or all process swatches
- Colors individual characters for single text frame
- Colors objects in order when multiple objects are selected
- Random shuffle if there are more than 3 colors
- Japanese and English UI support

### Process Flow

1. Check document and selection
2. Get swatches (if none selected, use all process swatches)
3. Apply colors per character for single text, or in order for multiple objects
4. Shuffle randomly if needed

### Original / Acknowledgements

Inspired by sort_by_position.jsx by shspage  
https://gist.github.com/shspage/02c6d8654cf6b3798b6c0b69d976a891

### Update History

- v1.0.0 (20241103): Initial version
- v1.1.0 (20250625): Supported all process swatches when no swatches are selected