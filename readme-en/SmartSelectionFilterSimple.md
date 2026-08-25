# SmartSelectionFilterSimple

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartSelectionFilterSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SmartSelectionFilterSimple.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSelectionFilterSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Filters the selection by condition and reselects text frames, open paths, or closed paths.

Switching the scope extends the filter from the top level of the selection to objects inside groups as well.

### Features

- Filter by text, open path, or closed path
- Switchable scope (selected objects only, or objects inside groups too)
- Closed paths include fill-only, stroke-only, and fill-plus-stroke paths

### Usage

1. Select the objects you want to filter.
2. Run the script.
3. Choose the condition and the scope, then confirm.

### Notes

- Compound paths are treated as parent objects.
- Clipping mask paths inside clipping groups are excluded from the targets.
- Selection updates avoid locked and hidden objects, and locked or hidden parent containers.
- Use SmartSelectionFilter.jsx when you need finer conditions.

### Update History

- v1.0
