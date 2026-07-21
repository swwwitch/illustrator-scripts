# ImportStyles

[![Direct](https://img.shields.io/badge/Direct%20Link-ImportStyles.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportStyles.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Picks a pre-registered AI file from a list and imports its contents into the current document.
- Graphic styles, brushes, symbols, and font samples come along with the pasted objects.
- The candidate list is persisted in a TSV file (`ImportStyles_candidates.tsv`) and can be extended from the dialog.

### Main Features

- Candidates are shown in a headered, scrollable 2-column ListBox (Content / Filename)
- **Category** radio buttons (**Style/Brush/Symbol** or **Fonts**) filter the list (no filter by default)
- Search field plus a **Search** button filters both content name and filename, case-insensitively (applied only on button click)
- Confirm with a double-click, the Enter key, or the **Load** button
- **Add** button: choose any AI file, copy it into the styles folder, then register or overwrite it in the candidates TSV with a display name and a category (auto-detected from the filename)
- Objects are pasted into the "// _imported" layer (created automatically if missing, and unlocked and made visible)
- The candidates TSV has three columns, `label / path / category`; the legacy check column (0|1) is ignored on load
- When no TSV exists, built-in defaults are used (Open path, Arrow, Circled numbers, Standard Latin fonts)
- Automatic Japanese / English UI

### Workflow

1. Choose an AI file from the candidate list in the dialog
2. Open it and copy every object inside the working artboard
3. Close the AI file without saving
4. Return to the original document, prepare the "// _imported" layer, and paste

### Not Supported

- No destination document open (shows an alert and exits)
- A candidate path that no longer exists (shows an alert and exits)
- Objects outside the working artboard are not copied
- Due to Illustrator's behavior, at least one Graphic Style / Brush always remains

### Update History

- v1.0 (2025-08-14): Initial release
- v1.1 (2025-08-15): Added delete option in CANDIDATES to remember after loading
- v1.2 (2025-08-15): Load CANDIDATES from an external TSV file
- v1.3 (2025-08-15): Added Category column (Style/Brush/Symbol/Font) and a category dropdown on Add
- v1.4 (2025-08-15): Switched to a standard 2-column ListBox, removed the delete option, added category radios and a Search button, unified the destination layer to "// _imported", refreshed docs
- v1.5.0 (2026-07-01): Structured localization (nested LABELS + dotted L()), wrapped everything in an IIFE, added shared panel setup (setupPanel), tidied variable/function names, split the add flow into functions, reduced duplication and unnecessary try blocks
