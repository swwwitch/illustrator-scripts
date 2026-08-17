# MakeRectangleFromGuides

[![Direct](https://img.shields.io/badge/Direct%20Link-MakeRectangleFromGuides.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/MakeRectangleFromGuides.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeRectangleFromGuides.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Fills every area bounded by vertical and horizontal guide intersections with a generated rectangle. A dialog selects which guides to use and where the rectangles go.

### Dialog

#### Source Guides

**Guide Type**

| Option | Description |
| --- | --- |
| All Guides | Uses every path marked as a guide (default) |
| Ruler Guides Only | Uses only the guides that run past the artboard edges |

**Layers**

| Option | Description |
| --- | --- |
| All Layers | Looks for guides across the whole document (default) |
| Current Layer Only | Looks only inside the active layer |

**Filters**

| Option | Description |
| --- | --- |
| Current Artboard Only | Uses only the guides that fall inside the current artboard (default: off) |
| Include Locked Layers | Temporarily unlocks locked layers to collect their guides, then restores the lock (default: on) |

#### Rectangles to Create

**Destination**

| Option | Description |
| --- | --- |
| Current Layer | Draws into the active layer (default). Stops if that layer is locked or hidden |
| Specific Layer | Draws into the layer named in the field. An existing layer with the same name is reused; otherwise it is created. Default name is `Generated Rectangles` |

**After Creation**

| Option | Description |
| --- | --- |
| Merge Into a Single Path | Unites every generated rectangle into a single path (default: off) |
| Convert to Shape | Converts the result into live shapes (default: on). When merging, the conversion runs after the merge |

The top of each panel reports how many guides matched (left) and how many rectangles will be created (right). Both refresh as the settings change, so the result can be checked before pressing OK.

### Process Flow

1. Collect the guides that match the settings (temporarily unlocking layers if needed)
2. Classify them as vertical or horizontal and collapse overlapping coordinates
3. Prepare the destination layer
4. Create one rectangle per bounded area
5. Merge and/or convert to live shapes if requested

### Notes

- Rectangles are filled with black in RGB documents and K100 otherwise, at 50% opacity.
- Locked sublayers are not covered by the temporary unlock.
- With "All Guides", a path such as a rectangle turned into a guide contributes only one of its edges.
- "Merge Into a Single Path" unites everything that is selected, so rectangles that are not touching also end up in one path.
- Fill color, opacity, and the default layer name can be changed in the "User settings" block at the top of the script.

### Update History

- v1.1.0 (20260817): Added Layers, Current Artboard Only, the Specific Layer destination, and Convert to Shape. The dialog now reports the guide and rectangle counts. Overlapping guides no longer produce zero-size rectangles
- v1.0 (20250713): Initial version

### Script info

- Version: v1.1.0
