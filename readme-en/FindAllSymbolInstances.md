# Select all instances of the same symbols

[![Direct](https://img.shields.io/badge/Direct%20Link-FindAllSymbolInstances.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/FindAllSymbolInstances.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FindAllSymbolInstances.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Finds every instance of the same symbols as the selected symbol instances across the document and reselects them all.

### Features

- Collects symbol instances nested inside groups recursively
- Handles selections that mix several symbols (searches per symbol and selects all the results together)
- When the selection contains no symbol instance, runs Select > Same > Appearance instead

### Usage

1. Select the symbol instances you want to use as the reference (selecting a group that contains them also works).
2. Run the script.

### Notes

- One instance is picked per symbol, Select > Same > Symbol Instance is run for each, and all the results are selected together.
- Items that cannot be selected (because they are locked, for example) are skipped.
- An alert appears when nothing could be selected.

### Article

https://note.com/dtp_tranist/n/n140952ad5011 (Japanese)

### Update History

- v1.1.1 (2026-09-15) Added a link to the article, localized the alert into Japanese and English, and reorganized internal naming and processing
- v1.1.0 (2026-05-09)
