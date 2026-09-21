# Apply an area-type graphic style

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyAreaTypeStyle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/single-function/ApplyAreaTypeStyle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyAreaTypeStyle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Pick a graphic style with the dialog's radio buttons (white text / frame only)
- If the chosen style is not in the document, it is imported from a predefined AI file (`TARGET_FILE_PATH`)
- The chosen graphic style is applied to the selected objects

### Settings

**This script will not work as shipped.** Edit the "User Settings" block at the top of the script to match your own environment before running it.

| Variable | Default | What it is |
| --- | --- | --- |
| `TARGET_FILE_PATH` | `/Users/takano/sw Dropbox/.../StyleForAreaType.ai` | **Absolute path** to the AI file the graphic styles are imported from. This is the author's own path, so you must replace it with the path to your own file |
| `STYLE_NAME_WHITE_TEXT` | `文字白抜き` | Name of the graphic style applied by the "White text" radio button |
| `STYLE_NAME_FRAME_ONLY` | `枠のみ` | Name of the graphic style applied by the "Frame only" radio button |

The style names have to match the names registered in the source AI file.

### Process Flow

1. Select the target objects and run the script
2. Choose the style (white text / frame only) in the dialog
3. If the style is not registered, the source AI is opened and the style is imported by copy and paste, then the temporary objects and layer are deleted
4. The graphic style is applied to the selected objects

### Notes

- During import the content is pasted onto a temporary `// _imported` layer, which is deleted together with its contents once the asset is registered (nothing changes visually)
- If the AI file at `TARGET_FILE_PATH` cannot be found, the style cannot be imported (see Settings)

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
