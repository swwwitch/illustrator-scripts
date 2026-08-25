# export-Event

[![Direct](https://img.shields.io/badge/Direct%20Link-export--Event.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/export-Event.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export-Event.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Exports every artboard of the active document to PNG, following per-name rules.

### Export rules

- `title` / `title2` (including names with a `-...` suffix): exported at 100% on a transparent background
- `Doorkeeper`: exported at 100% and 200% on a white background (the 200% file gets a `-200` suffix)
- `シンボル一覧` (Symbol List): excluded from the export
- Anything else: exported at 100% on a white background

Edit `buildExportJobs()` to add or change rules. Returning an empty array excludes the artboard; returning several entries exports it at several scales.

### Script info

- Version: v1.0.4
