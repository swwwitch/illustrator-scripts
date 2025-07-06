# SmartLayerManage.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartLayerManage.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/SmartLayerManage.jsx)


[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script to batch move objects to a specified layer.
- Supports switching modes: selected objects, all text, all objects, or all (force).

### Main Features

- Mode selection (Selected / Text Only / All / All (Force))
- Option to delete empty layers (excluding layers starting with bg or //)
- Automatically change target layer color to RGB(79,128,255)
- Unlocking, showing, and recursive item collection
- Japanese and English UI support

### Process Flow

1. Select mode and target layer in the dialog
2. Collect target objects
3. Move objects to the selected layer
4. Optionally delete empty layers
5. Change target layer color

### Update History

- v1.0.0 (20250703): Initial release
- v1.0.1 (20250703): Added layer color change function
- v1.0.2 (20250703): Improved auto selection detection and empty layer deletion logic
- v1.0.3 (20250704): Added "All (Force)" mode (merge all layers)