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
- Layer Name Format
  - Toggle "Include number" and "Include name" (they cannot both be off)
  - Separator: Underscore (_) / Hyphen (-) / Space / None (default: underscore)
  - The separator applies only when both the number and the name are included
  - The resulting name is previewed live in the dialog, e.g. "Example: 1_Cover"
  - Falls back to "Artboard" when the artboard name is empty
- Exclude options
  - Locked: Layer / Object, Hidden: Layer / Object (all on by default)
  - Whatever you turn off is temporarily unlocked / shown for the move, then restored to its original state
  - "Specify by name" accepts comma-separated layer names to leave untouched (default: bg)
  - Guides bypass the name-based exclusion and are still collected on _guide (locked / hidden ancestors are respected)
- "After Organizing" toggle removes empty layers and sub-layers (on by default)
  - Sub-layers matching an exclusion rule are kept together with their contents
- Legacy layers (named after the artboard only) are merged into the new layers and removed
- Target layers that are locked or hidden are temporarily unlocked / shown for the move, then restored
- "_guide" and "_pasteboard" are protected layers that are never deleted even when empty, including when they appear as sub-layers; _guide is brought to the front
- Reports the number of objects that could not be moved via an alert
- Automatic Japanese / English UI

### Workflow

1. Choose target artboards, layer name format, exclusions, and after-organizing options in the dialog
2. Collect the top-level objects and compute every centroid up front
3. Temporarily release the lock / hidden state of whatever is in scope, then move each object into the layer of the artboard containing its centroid
4. Send objects that belong to no artboard to _pasteboard / _guide and merge legacy artboard-named layers
5. Reorder the layers to match artboard order and restore the original lock / hidden state
6. Remove empty layers / sub-layers when requested

### Not Supported

- No open document (an alert is shown and the script exits)
- Locked or hidden layers and objects when the corresponding exclusions are on, plus any object under such an ancestor
- Layers listed under "Specify by name" and their contents (guides excepted)
- Objects nested inside groups or symbols (only top-level objects are processed)
- The _pasteboard pass and legacy-layer merge are skipped in "Current artboard" mode
- Protected layers (_guide / _pasteboard) are never deleted, even when empty

### Update History

- v1.3.1 (2026-08-17): Fixed objects being counted as failures instead of moved when the Locked / Hidden exclusions were turned off. Fixed the script aborting while removing empty layers, and failures being counted twice. Excluded sub-layers are now kept together with their contents. Added the layer name preview, revised the UI wording, and cached centroid calculation for speed
- v1.3.0 (2026-05-26)
