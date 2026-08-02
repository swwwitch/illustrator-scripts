# VariableDataImport.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-VariableDataImport.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/data/VariableDataImport.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VariableDataImport.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

A data-merge script for Illustrator that imports a CSV / TSV file into a template.

It replaces `<header>` placeholder tags inside text frames with row values and generates one artboard variation per data row — useful for business cards, place cards, membership cards, price tags, certificates, and anything else that shares a design but differs in content.

The original document is duplicated with Save As and the merge happens in the copy, so the original file is never modified.

## Main features

- Choose a CSV / TSV file from the same folder as the open document
- Review the imported data in a list inside the dialog
- Replace `<header>` tags with row values, preserving the character formatting applied to the tag
- Generate one artboard per data row automatically
- Pick the column count that makes the grid closest to a square, centred on the canvas
- Set the gap between artboards (rounded to 10 pt increments)
- Choose the data column used for artboard names
- Preview the result in a duplicate file without touching the original
- Reads UTF-8 (with or without BOM) and Shift-JIS data files
- Japanese / English UI

## Usage

1. Build the template design on **artboard 1**. Write the variable parts as tags inside text frames, e.g. `<name>` and `<company>`.
2. Put the data file (CSV / tab-separated text) in the **same folder** as the document. The first line is the header row, and its column names map to the tag names.
3. Save the document (the script cannot run on an unsaved document).
4. Run `VariableDataImport.jsx`.
5. Configure the dialog.
   - **File**: the data file to import
   - **Artboard name column**: the column used to name each artboard
   - **Artboard gap**: the spacing between duplicated artboards (the same horizontally and vertically, in points)
6. Turn on **Preview** to check the result in a duplicate file. Turning it off closes the preview.
7. Click **Duplicate and Run**. The original document is duplicated with Save As and the data is merged into the copy.

## Data file format

| Extension | Delimiter |
| --- | --- |
| `.csv` | Comma (quoted fields and `""` escapes are supported) |
| `.txt` | Tab |

```
name,company,title
Taro Yamada,Switch Inc.,CEO
Hanako Suzuki,Sample Inc.,Designer
```

With this data, text frames containing `<name>`, `<company>`, and `<title>` are replaced with the corresponding values.

- Header names in the first line become the tag names.
- Surrounding whitespace is stripped from every value.
- Rows with an empty name value fall back to `Data_1`, `Data_2`, and so on.

## Generated files

| Action | Generated file |
| --- | --- |
| Duplicate and Run | `<original name>_import_YYYYMMDD_HHMMSS.ai` |
| Preview | `<original name>_preview_YYYYMMDD_HHMMSS.ai` (removed when Preview is turned off) |

Both are created next to the original document, which stays unchanged.

## Notes

- Only **unlocked and visible** objects on artboard 1 are duplicated as the template. Backgrounds on locked layers are not carried into the variations.
- Preview copies the file on disk, so unsaved edits are not reflected.
- Newlines inside quoted fields are not supported; write one record per line.
- Up to 1000 artboards can be generated. If the rows do not fit on the canvas, reduce the artboard gap.
- The first row is merged last, so the template tags survive until every other row is done.
- Tags are replaced in text frames and in text frames inside groups. Text inside symbols or embedded images is not processed.

## Article

[Illustrator: a script for merging CSV data | DTP Transit](https://note.com/dtp_tranist/n/n741c9f28d0fd)

## Changelog

- v1.4.1 (2026-05-18): Stability improvements
- v1.4.0 (2026-05-15): Added Preview
- v1.3 (2026-01-22): Switched artboard placement to an automatic grid calculation
