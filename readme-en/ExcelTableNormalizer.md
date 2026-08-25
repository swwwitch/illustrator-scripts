# ExcelTableNormalizer

[![Direct](https://img.shields.io/badge/Direct%20Link-ExcelTableNormalizer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/ExcelTableNormalizer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExcelTableNormalizer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Normalizes Illustrator artwork that came from Excel so that it is easier to work with as a table.

### Features

- Releases clipping masks and deletes small stray objects such as error indicators
- Moves text to a dedicated layer (`_text_all`) and removes duplicate text
- Per-column alignment chosen in a dialog (left / center / right, with auto-detection)
- Converts text to a print-ready black
- Infers columns from the rule grid and shows column samples
- Extracts and adjusts cell backgrounds on a dedicated layer (`_cell_rectangle`), equalizing heights when needed
- Converts rectangular rules to center lines, with even distribution, merged-cell support and column-width preservation

### Usage

1. Open the document containing the table pasted from Excel.
2. Run the script.
3. Set the per-column alignment and run it.

### Notes

- Color-mode menu commands are run at startup so that later color assignments behave consistently.
- The script rebuilds the artwork substantially, so duplicating the file beforehand is recommended.

### Update History

- v1.1.0 (2026-04-30)
