# CollectGuides

[![Direct](https://img.shields.io/badge/Direct%20Link-CollectGuides.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/CollectGuides.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CollectGuides.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Collects guides scattered across multiple layers and sub-layers into a single layer (`// guide` by default).
- Hidden and locked layers and guides are included: they are unlocked and shown during the run and restored afterwards.

### Main Features

- Walks every layer and sub-layer recursively and moves the guides
- Saves and restores the locked, visible and hidden state of layers and items
- The target layer name is configurable via `TARGET_GUIDE_LAYER_NAME`
- Optionally restores the global guide visibility and lock state afterwards
- For large documents, scans only the relevant layers and suppresses redraws

### Usage

1. Open the target document
2. Run the script

### Notes

- Show Guides and Unlock Guides are invoked temporarily during the run.
- An existing target layer is reused.

### Update History

- v1.0 (20250816): Initial version with guide collection, recursive walking, and lock/visibility restore
- v1.1 (20250816): Added a narrowed scan that only visits layers containing guides
- v1.2 (20250816): Added settings for restoring guide visibility and lock afterwards
- v1.3 (20250816): Made the target layer name configurable
- v1.4 (20250816): Added a redraw-suppression option
