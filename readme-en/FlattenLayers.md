# FlattenLayers

[![Direct](https://img.shields.io/badge/Direct%20Link-FlattenLayers.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenLayers.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Created

2025-04-14

### Updated

2026-04-15

### Overview

- Show a dialog before execution so the processing conditions can be selected
- Flatten objects from non-excluded layers and sublayers into the specified destination layer
- Let the user choose whether locked / hidden layers and objects are excluded
- Automatically dim UI items when there are no locked layers, hidden layers, hidden objects, guides, or sublayers in the document
- Toggle all exclusion options on / off only for the items that are currently enabled
- Optionally move remaining non-empty sublayers to the top level
- Let the user choose how guides are handled: integrate, keep in the current layer, or move to another layer
- In “Keep in the current layer” mode, guides directly under sublayers are hoisted to the parent layer
- In “Move to another layer” mode, guides are moved to the specified guide layer after flattening
- Delete empty layers / sublayers repeatedly until no more removable empty layers remain
- Apply skipLockedLayers / skipHiddenLayers consistently to both top-level layers and sublayers
- Layers that still contain only guides are not treated as empty
- Let the user specify the destination layer name, reuse an existing destination layer, and set the destination layer color
- If destination-layer reuse is off and the same name already exists, create a new layer with a numbered unique name

### Process Flow

1. Get the active document (exit if none)
2. Show the options dialog, automatically dim irrelevant items, and exit if canceled
3. Reuse the existing destination layer or create a new uniquely named one according to the selected options
4. Flatten items from eligible non-excluded layers and sublayers into the destination layer
5. If the guide mode is “Keep in the current layer,” hoist guides directly under sublayers to the parent layer
6. Optionally move remaining non-empty sublayers to the top level
7. If the guide mode is “Move to another layer,” move guides from the flattened result into the specified guide layer
8. If enabled, repeatedly delete empty layers / sublayers until none remain
9. If any failures occurred, show only the failure counts after processing

### Update History

- v1.0 (20250414) : Initial release
- v1.7.1 (20260408) : Added automatic dimming of irrelevant exclusion UI, enabled-only toggle-all behavior, consistent skipLockedLayers / skipHiddenLayers handling for sublayers, clearer guides-panel initialization, and updated overview/comments
- v1.7.3 (20260413) : Narrowed excluded layer names to only "bg" (removed "背景" and "background")
- v1.7.4 (20260415) : Added option to move guides from excluded layers to the guide layer (available when "Move to another layer" is selected, default ON)

### Script info

- Version: v1.7.4
