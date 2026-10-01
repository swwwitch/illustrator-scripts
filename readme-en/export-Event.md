# Export every artboard to PNG by naming rule

[![Direct](https://img.shields.io/badge/Direct%20Link-export--Event.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/export-Event.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export-Event.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Exports the artboards chosen in a dialog to PNG, following per-name rules.
The list shows each artboard's scale and background (transparent or white). Every artboard starts selected.
**Close the document after export** (on by default) closes the document when the export finishes; if there are unsaved edits, you are asked whether to save.
**Open the output folder after export** (on by default, macOS only) opens the output folder when the export finishes.

### Export rules

- `title` / `title2` (including names with a `-...` suffix): exported at 100% on a transparent background
- `Doorkeeper`: exported at 100% and 200% on a white background (the 200% file gets a `-200` suffix)
- `シンボル一覧` (Symbol List): excluded from the export
- Anything else: exported at 100% on a white background

Edit `buildExportJobs()` to add or change rules. Returning an empty array excludes the artboard; returning several entries exports it at several scales.

### Opening in Path Finder

With the helper app `/Applications/OpenInFileViewer.app` installed, the output folder opens in Path Finder when it is running, and in Finder otherwise. Without the helper app, it opens in Finder.

Build the helper app from [helpers/OpenInFileViewer.applescript](https://github.com/swwwitch/illustrator-scripts/blob/master/helpers/OpenInFileViewer.applescript) in this repository:

```
osacompile -o /Applications/OpenInFileViewer.app helpers/OpenInFileViewer.applescript
```

### Update History

- v1.2.0 (2026-10-01) Added **Close the document after export** and **Open the output folder after export** below the list (both on by default; opening the folder is macOS only). With the helper app, the folder opens in Path Finder while it is running
- v1.1.2 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.1.1 (2026-10-01) The progress palette's Cancel button now uses the shared button-row part. Unified the window and panel margins and spacing with the shared layout part
- v1.1.0 (2026-09-30) Choose the artboards to export in a dialog; the list shows the scale and background
- v1.0.5 (2026-09-27) Alerts and the progress window now switch between Japanese and English

### Script info

- Version: v1.2.0
