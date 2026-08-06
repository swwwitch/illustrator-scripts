# RenameArtboardsPlus.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-RenameArtboardsPlus.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/RenameArtboardsPlus.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script to flexibly batch rename artboards in Illustrator.
- Combine prefixes, suffixes, file name, numbers, and original artboard names to create advanced naming rules.

### Main Features

- Freely choose whether to include prefix, suffix, or original artboard name
- Supports file name reference, separator settings, preset save/load
- Numbering formats: number, uppercase alphabet, lowercase alphabet
- Auto-detect padding digits based on start number (zero-padding)
- Preview feature to check the renaming result beforehand
- Displays up to 15 preview items, improved button layout
- Compatible with ExtendScript (ES3)

### Process Flow

1. Configure options (prefix, suffix, numbering format, etc.) in dialog
2. Check name examples with the preview
3. Click "OK" to rename artboards

### Update History

- v1.3.1 (2026-08-06): Write exported presets as UTF-8 so Japanese labels are not garbled, dim the suffix separator when the numbering format is "none" (it has no effect there), reject non-digit input and ignore surrounding whitespace in start number / increment, keep the dialog open when renaming fails, unified panel margins and spacing through a shared layout setup (setupPanel / PANEL_MARGINS), renamed variables / functions / panels to match what they actually represent, grouped LABELS into categories, and added JSDoc to every function
- v1.3.0 (2026-05-09): Added a warning when the resulting name would be empty, tagged separator radio buttons with their values to remove array-order dependency, escaped strings when exporting presets, changed the default start number to "001"
- v1.2.0 (2026-05-07): Localization overhaul (L / labelText / labelWithCount), key-based internals for EN locale support, added "none" numbering format, removed the Apply button, split main() into UI builders / preset I/O / pure helpers
- v1.1 (2025-04-30): Added auto-detection of padding digits, simplified preset labels, enhanced ES3 support
- v1.0 (2025-04-20): Initial version created