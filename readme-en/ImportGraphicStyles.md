# ImportGraphicStyles

[![Direct](https://img.shields.io/badge/Direct%20Link-ImportGraphicStyles.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportGraphicStyles.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportGraphicStyles.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- The Load Styles button picks an AI file and imports the graphic styles defined in it
- Radio buttons are generated automatically from the imported style names, and the chosen style is applied to the selection
- The chosen file is remembered and reused on later runs, until another file is picked with the button

### Process Flow

1. Select the target objects and run the script
2. If a file is remembered, its styles are imported and shown as radio buttons
3. Picking another file with Load Styles re-imports it and refreshes the radios
4. Choose a style and click Apply to apply it to the selected objects

### Notes

- During import the content is pasted onto a temporary `// _imported` layer, which is deleted together with its contents once the asset is registered (nothing changes visually)
- The source file path is stored in `Folder.userData` (`swwwitch_ImportGraphicStyles.txt`), not inside the script
- The radios are generated from the graphic style names in the source file; the default style at index 0 is excluded

### Update History

- v1.7.0 (20260701): Replaced the hard-coded source file with a Load Styles button (remembered in `Folder.userData`), generated the radios from the imported style names, and removed the fixed style names
- v1.6.0 (20260701): Removed search, add and the candidate list; the script now imports the style chosen by radio button when needed and applies it to the selection
- v1.5.0 (20260701): Structured the localization, wrapped everything in an IIFE, added a shared panel helper, tidied names, split the import flow into functions, and removed duplicated code
- v1.4 (20250815): Switched to a standard ListBox, dropped the delete option, added category radios and a search button, and unified the paste target
- v1.3 (20250815): Added a category column and a category dropdown when adding
- v1.2 (20250815): Allowed CANDIDATES to be loaded from an external TSV
- v1.1 (20250815): Recorded the delete option in CANDIDATES

### Script info

- Version: v1.7.0
