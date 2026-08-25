# FindAllSymbolInstances

[![Direct](https://img.shields.io/badge/Direct%20Link-FindAllSymbolInstances.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/FindAllSymbolInstances.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FindAllSymbolInstances.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Finds every instance of the same symbols as the currently selected symbol instances and reselects them all.

### Features

- Collects symbol instances nested inside groups recursively
- Handles selections that mix several different symbols
- Falls back to searching by appearance when the selection contains no symbol instance

### Usage

1. Select the symbol instances you want to use as the reference.
2. Run the script.

### Notes

- One representative is picked per symbol definition, the Select Symbol Instance command is run for each, and the results are merged without duplicates.
- Items that cannot be reselected (because they are locked, for example) are skipped silently.
- An alert appears only when nothing could be collected.

### Update History

- v1.1.0 (2026-05-09)
