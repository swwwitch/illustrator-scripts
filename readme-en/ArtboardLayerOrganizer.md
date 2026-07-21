# ArtboardLayerOrganizer

[![Direct](https://img.shields.io/badge/Direct%20Link-ArtboardLayerOrganizer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/ArtboardLayerOrganizer.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that distributes the objects in a document by artboard and organizes them into layers named "number_artboard name".
- Each object is assigned to an artboard by its centroid.
- Guides are collected on "_guide" and objects that belong to no artboard on "_pasteboard" (both created only when needed).
- Layer order is aligned with artboard order (1 -> 2 -> 3 ... from the top).

### Main Features

- Target Artboards
  - **Current artboard** / **All** (locked to "Current artboard" when the document has only one artboard)
- Layer Name composition
  - Toggle "Artboard Number" and "Artboard Name"
  - Separator: Underscore (_) / Hyphen (-) / Space / None (default: underscore)
  - Falls back to "Artboard" when the artboard name is empty
- Exclude options
  - Locked: Layer / Object, Hidden: Layer / Object (all on by default)
  - "Specified Layers" accepts comma-separated layer names to leave untouched (default: bg)
  - Guides bypass the name-based exclusion and are still collected on _guide (locked / hidden ancestors are respected)
- Post-Processing toggle removes empty layers and sub-layers (on by default)
- Legacy layers (named after the artboard only) are merged into the new layers and removed
- Target layers that are locked or hidden are temporarily unlocked / shown for the move, then restored
- "_guide" and "_pasteboard" are protected layers that are never deleted even when empty; _guide is brought to the front
- Reports the number of objects that could not be moved via an alert
- Automatic Japanese / English UI

### Workflow

1. Choose target artboards, layer naming, exclusions, and post-processing in the dialog
2. Collect the top-level objects, resolve the owning artboard from each centroid, and move them into the matching layer
3. Send leftover objects to _pasteboard / _guide and merge legacy artboard-named layers
4. Reorder the layers to match artboard order and remove empty layers / sub-layers when requested

### Not Supported

- No open document (an alert is shown and the script exits)
- Locked or hidden layers and objects when the corresponding exclusions are on, plus any object under such an ancestor
- Layers listed in "Specified Layers" and their contents (guides excepted)
- The _pasteboard pass and legacy-layer merge are skipped in "Current artboard" mode
- Protected layers (_guide / _pasteboard) are never deleted, even when empty

### Update History

- v1.3.0 (20260526): Current version
