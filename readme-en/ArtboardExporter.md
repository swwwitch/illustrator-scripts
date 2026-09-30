# Export chosen artboards as PNG, JPEG or PDF

[![Direct](https://img.shields.io/badge/Direct%20Link-ArtboardExporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/ArtboardExporter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardExporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Exports the artboards chosen in a dialog as PNG, JPEG or PDF.
Each artboard can have its own format, scales (one or more) and background (white, black or transparent); file names, destination and exclusions are configurable too.

### Artboards to export

- The list shows the number, whether the artboard is exported (✓ / Excluded), the name, the format, the scale and the background
- All, Active Artboard and By Number (e.g. `1-3, 5`) switch the export targets at once
- Double-clicking a row also switches it
- Exclude Artboards: artboards whose names start with this are skipped (default `#`)
- Exclude Layers: layers whose names start with this are hidden while exporting and shown again afterwards (default `//`)

### Export Settings for Selected Artboards

Select rows in the list and change them here; selecting several rows changes them together. Every row starts selected.

- Export: rows turned off are skipped
- Format: PNG / JPEG / PDF (the last format chosen becomes the next initial format)
- Scale: separate values with commas, e.g. `100, 200`, to export at several scales. JPEG goes up to 776%
- Background: white, black or transparent (clicking a swatch also works; transparent is PNG only)
- Quality: JPEG quality (0-100; initial value from `JPEG_QUALITY`)
- PDF Preset: chosen for PDF (the last preset chosen becomes the next initial preset)
- Combine PDFs into One File: puts every PDF into one file (a global setting; uses the first artboard's preset)

### File Name

- Combine: File Name (editable) + the artboard part (number and/or name with a separator, or none) + a separator (`_` / `-`) + Date (YYYYMMDD / YYMMDD / YYYY-MM-DD / MMDD) + Text
- Pattern: write it freely, e.g. `{doc}_{name}_{date}`. Tokens: {doc} (document name), {name} (artboard name), {num} (number), {scale} (scale), {date} (date)
- Replace: characters file names cannot hold (`/ \ : * ? " < > | ¥`) become `_` or `-`; turn on Spaces to replace spaces too
- The panel shows an example file name

### Destination

- Same as Source File, or Custom (pick the folder with Choose…). Documents that were never saved can only use Custom
- Create Folder: exports into a new folder inside the destination ({doc} and {date} work)
- Show Folder After Export: opens the destination folder when the export finishes

### Formats

- PNG: uses the scale and background
- JPEG: uses the scale, the background (white or black) and the quality
- PDF: one file per artboard (or one combined file), saved with the chosen PDF preset; scale and background are not used. It is made from the saved file, so you are asked to save when there are unsaved changes (documents that were never saved cannot export PDF)

### Output

- File name (Combine): `[file name][artboard part][separator scale][separator date][separator text].<extension>` (the scale is added only for anything other than 100%)
- File name (Pattern): the tokens replaced; without {scale}, `-scale` is appended for anything other than 100%
- Combined PDF: named without the artboard part and scale
- Numbers are zero-padded to the digit count of the artboards (01, 02 … with ten or more)
- When file names would collide, nothing is exported and you are told which ones
- When files with the same names exist, you are asked once whether to overwrite them

### Memory

Clicking Export remembers each artboard's settings plus the file name, destination, exclusion and other settings until Illustrator quits. The next time the dialog opens for the same document, artboards with the same names get their settings back. An edited File Name field is not remembered.

### Update History

- v1.0.2 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.0.1 (2026-10-01) The progress palette's Cancel button now uses the shared button-row part. Unified the window and panel margins and spacing with the shared layout part
- v1.0.0 (2026-09-30) First release

### Script info

- Version: v1.0.0
