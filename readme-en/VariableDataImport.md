# VariableDataImport.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-VariableDataImport.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/data/VariableDataImport.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VariableDataImport.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

A data-merge script for Illustrator that imports a CSV / TSV file into a template.

It replaces `<tag>` placeholders inside text frames with row values and generates one artboard variation per data row — useful for business cards, place cards, membership cards, price tags, certificates, and anything else that shares a design but differs in content.

The original document is duplicated with Save As and the merge happens in the copy, so the original file is never modified.

## Main features

- Choose a CSV / TSV file from the same folder as the open document
- Review the imported data in a list inside the dialog
- Replace `<tag>` placeholders with row values, preserving the character formatting applied to the tag
- Match variables to columns by column name, or map them manually
- In manual mode, pick the data column for each variable (columns with a matching name are selected automatically)
- Each popup menu shows the value of the first data row (line 2 of the file) beside it
- Columns already used by another variable are dimmed in the popup menu
- Derive the amount excluding tax and the tax itself from a tax-included column
- Generate one artboard per data row automatically
- Pick the column count that makes the grid closest to a square, or set the count explicitly
- Set the gap between artboards (rounded to 10 pt increments)
- Step numeric fields with the arrow keys (shift for tens, option for tenths)
- Choose the data column used for artboard names (a column containing 名前 / 御中 / 宛先 / 様 / 会社名 is preselected)
- Configure the saved file name (base name, suffix, separator, date, time)
- Preview the result in a duplicate file without touching the original
- Reads UTF-8 (with or without BOM) and Shift-JIS data files
- Japanese / English UI

## Usage

1. Build the template design on **artboard 1**. Write the variable parts as tags inside text frames, e.g. `<name>` and `<company>`.
2. Put the data file (CSV / tab-separated text) in the **same folder** as the document. The first line is the header row. Columns whose name matches a tag are mapped automatically, and any other column can be picked in the dialog.
3. Save the document (the script cannot run on an unsaved document).
4. Run `VariableDataImport.jsx`.
5. Configure the dialog.
   - **Match by column name** / **Map manually**: how the mapping is decided. Matching by column name uses the data column whose name matches each `<tag>` as-is. Choosing manual reveals the **Variable Mapping** panel.
   - **File**: the data file to import
   - **Variable Mapping**: the data column merged into each `<tag>` on the canvas (columns with a matching name are selected automatically; choose **(none)** to leave that tag untouched)
   - **Artboard name**: the data column used to name each artboard
   - **Artboard gap**: the spacing between duplicated artboards (the same horizontally and vertically, in points)
   - **Columns**: how many artboards to place side by side. Leave it empty to pick the count that makes the grid closest to a square.
   - **File Name**: the name of the duplicated file
6. Turn on **Preview** to check the result in a duplicate file. Turning it off closes the preview.
7. Click **Duplicate and Merge**. The original document is duplicated with Save As and the data is merged into the copy.

### Variable mapping

Choosing **Map manually** lists every `<tag>` found in the text on artboard 1.

- Beside each popup menu is the value of the **first data row (line 2 of the file)** in the selected column, so you can confirm the choice against real data.
- A column already used by another variable is dimmed and cannot be picked again.
- Changing a mapping turns the preview off. Turn **Preview** back on to rebuild it with the new mapping.

### Consumption tax

When the data only carries a tax-included price, the amount excluding tax and the tax itself can be derived. Turn on **Calculate consumption tax** in the **Variable Mapping** panel and the derived items appear in the popup menus.

| Item | Value |
| --- | --- |
| `price` | The value as it appears in the data |
| `price (excl. tax)` | `price / 1.1`, rounded |
| `price (tax)` | `price - (excl. tax)` |

The two derived values always add back up to the original amount. Thousands separators, full-width digits, and a `¥` sign are all accepted, and a value that came in with separators keeps them.

The derived items appear only when a single tax-included column can be identified. A column whose name contains 税込 / 金額 / 価格 / 料金 / 定価 / 合計 / price / amount / total / cost / fee wins; otherwise the column is used only when exactly one numeric column exists. When several candidates compete, the checkbox is disabled.

### File name

The name of the duplicated file is assembled from these parts.

| Item | Meaning |
| --- | --- |
| Base name | The body of the name. Leave it empty to reuse the original file name |
| Suffix | Inserted after the base name, joined with the separator. Empty by default, so nothing is inserted |
| Append date | `YYYYMMDD` at the end |
| Append time | `HHMMSS` at the end |
| Separator | The character joining the parts (hyphen or underscore) |

The assembled name is shown under **Saved as**. `\ / : * ? " < > |` and surrounding whitespace are stripped.

### Keyboard stepping

**Artboard gap** and **Columns** accept keyboard stepping.

| Key | Step |
| --- | --- |
| Up / Down | ±1 |
| Shift + Up / Down | Snap to the nearest ten |
| Option + Up / Down | ±0.1 |

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

- A header name identical to a tag name is mapped automatically; otherwise pick the column under **Variable Mapping** in the dialog.
- Surrounding whitespace is stripped from every value.
- Rows with an empty name value fall back to `Data_1`, `Data_2`, and so on.

## Generated files

| Action | Generated file |
| --- | --- |
| Duplicate and Merge | Whatever the **File Name** panel specifies (by default `<original name>_YYYYMMDD_HHMMSS.ai`) |
| Preview | `<base name>_preview_YYYYMMDD_HHMMSS.ai` (removed when Preview is turned off) |

Both are created next to the original document, which stays unchanged. The preview file always carries a date and time regardless of the **Append date** / **Append time** settings, so it never collides with the saved name.

## Notes

- Only **unlocked and visible** objects on artboard 1 are duplicated as the template. Backgrounds on locked layers are not carried into the variations.
- Preview copies the file on disk, so unsaved edits are not reflected.
- Newlines inside quoted fields are not supported; write one record per line.
- Up to 1000 artboards can be generated. If the rows do not fit on the canvas, reduce the artboard gap.
- The first row is merged last, so the template tags survive until every other row is done.
- Tags are replaced in text frames and in text frames inside groups. Text inside symbols or embedded images is not processed.
- The dialog lists the `<...>` tags found in the text on artboard 1, so a tag that is not written in the template does not appear.
- The tax rate is fixed at 10%.
- If the requested **Columns** value leaves too many rows to fit, the "too much data" warning appears. Clear the field to fall back to the automatic layout.

## Article

[Illustrator: a script for merging CSV data | DTP Transit](https://note.com/dtp_tranist/n/n741c9f28d0fd)

## Changelog

- v1.5.0 (2026-08-25): Added match-by-column-name and manual mapping modes, sample values and dimming of used columns in the mapping dropdowns, consumption-tax derivation, an explicit column count, file-name settings, and keyboard stepping for numeric fields
- v1.4.1 (2026-05-18): Stability improvements
- v1.4.0 (2026-05-15): Added Preview
- v1.3 (2026-01-22): Switched artboard placement to an automatic grid calculation
