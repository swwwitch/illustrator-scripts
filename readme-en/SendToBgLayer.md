# SendToBgLayer.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-SendToBgLayer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/SendToBgLayer.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script to move selected objects to a "bg" layer while preserving their stacking order and placing them at the very back.
- Automatically creates a "bg" layer if it does not exist, and locks it after processing.

### Main Features

- Move selected objects to "bg" layer
- Preserve stacking order
- Auto-create "bg" layer if missing
- Send "bg" layer to back and lock it
- Restore original active layer

### Process Flow

1. Get document and active layer
2. Check if "bg" layer exists, create if not
3. Move selected objects preserving stacking order
4. Send "bg" layer to back and lock
5. Reactivate original layer

### Update History

- v1.0.0 (20240624): Initial version