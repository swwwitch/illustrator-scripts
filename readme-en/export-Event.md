# Export every artboard to PNG by naming rule

[![Direct](https://img.shields.io/badge/Direct%20Link-export--Event.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/export-Event.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export-Event.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Exports the artboards chosen in a dialog to PNG, following per-name rules.
The list shows each artboard's scale and background (transparent or white). Every artboard starts selected.

### Export rules

- `title` / `title2` (including names with a `-...` suffix): exported at 100% on a transparent background
- `Doorkeeper`: exported at 100% and 200% on a white background (the 200% file gets a `-200` suffix)
- `シンボル一覧` (Symbol List): excluded from the export
- Anything else: exported at 100% on a white background

Edit `buildExportJobs()` to add or change rules. Returning an empty array excludes the artboard; returning several entries exports it at several scales.

### Update History

- v1.1.0 (2026-09-30) Choose the artboards to export in a dialog; the list shows the scale and background
- v1.0.5 (2026-09-27) Alerts and the progress window now switch between Japanese and English

### Script info

- Version: v1.1.0
