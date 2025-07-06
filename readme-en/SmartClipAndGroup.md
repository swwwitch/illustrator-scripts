# SmartClipAndGroup.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartClipAndGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/SmartClipAndGroup.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script to group selected objects based on overlap ratio or distance threshold, or to create clipping masks using the topmost or bottommost object as a reference.
- Supports flexible mask processing including placed-image-only clipping and square masks.

### Main Features

- Grouping by overlap ratio or proximity
- Clipping masks using topmost or bottommost object
- Placed-image-only rectangular or square masks
- Restore initial threshold value on re-execution
- Japanese and English UI support

### Process Flow

1. Choose processing mode (grouping or clipping) and threshold in dialog
2. Execute grouping or mask processing based on user choice
3. Update selection with resulting objects

### Update History

- v0.0.1 (20240605): Initial release
- v0.0.2 (20240610): Simplified UI and restructured
- v0.0.3 (20240610): Added placed-only and square mask options
- v0.0.4 (20240610): Added overlap-based grouping
- v0.0.5 (20240610): Improved z-order retention, re-execution support, and threshold restore