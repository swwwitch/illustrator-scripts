# ImportGraphicStyles-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-ImportGraphicStyles-v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportGraphicStyles-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportGraphicStyles-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Pick a graphic style with the dialog's radio buttons (white text / frame only)
- If the chosen style is not in the document, it is imported from a predefined AI file (`TARGET_FILE_PATH`)
- The chosen graphic style is applied to the selected objects

### Process Flow

1. Select the target objects and run the script
2. Choose the style (white text / frame only) in the dialog
3. If the style is not registered, the source AI is opened and the style is imported by copy and paste, then the temporary objects and layer are deleted
4. The graphic style is applied to the selected objects

### Notes

- During import the content is pasted onto a temporary `// _imported` layer, which is deleted together with its contents once the asset is registered (nothing changes visually)
- The source file is set by `TARGET_FILE_PATH`, and the style names by `STYLE_NAME_WHITE_TEXT` / `STYLE_NAME_FRAME_ONLY`

### Update History

- v1.6.0 (20260701): Removed search, add and the candidate list (ListBox / category selection). The script now imports the graphic style chosen by radio button (white text / frame only) when needed and applies it to the selection
- v1.5.0 (20260701): Structured the localization (nested LABELS plus a dotted lookup), wrapped everything in an IIFE, added a shared panel helper, tidied variable and function names, split the import flow into functions, and removed duplicated code and unnecessary `try` blocks
- v1.4 (20250815): Switched to a standard ListBox (two-column header), dropped the delete option, added category radios and a search button, unified the paste target to `// _imported`, and refreshed the documentation
- v1.3 (20250815): Added a category column (style / brush / symbol / font) and a category dropdown when adding
- v1.2 (20250815): Allowed CANDIDATES to be loaded from an external TSV
- v1.1 (20250815): Recorded the delete option in CANDIDATES
- v1.0 (20250814): Initial version

### Script info

- Version: v1.6.0
