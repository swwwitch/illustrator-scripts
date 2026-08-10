# RenameAssets

[![Direct](https://img.shields.io/badge/Direct%20Link-RenameAssets.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/RenameAssets.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RenameAssets.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Renames graphic styles, brushes, swatches and symbols in bulk using a find-and-replace pair entered in the dialog.
- A preview (triggered by the button, not while typing) shows the result beforehand, and Cancel reverts it.

### Main Features

- Find and replace, with regular expression and ignore-case options
- Switch the target collection (styles / brushes / swatches / symbols) with radio buttons
- For graphic styles only, a multi-select filter narrows the styles to rename
- Duplicate names are avoided automatically by appending (2), (3), and so on
- Default names wrapped in square brackets are skipped
- Dialog opacity and position can be adjusted
- Japanese and English localization

### Process Flow

1. Choose the target collection, the find and replace strings, and the options (regex, case) in the dialog
2. For styles, narrow the target styles with the multi-select filter if needed
3. Press Preview to apply the names to the panel temporarily (safely restorable from a snapshot)
4. Click OK to commit, or Cancel to revert

### Notes

- There is no live preview while typing; the preview runs only when the button is pressed.

### Update History

- v1.0 (20250820): Initial version

### Script info

- Version: v1.0
