# Open files, save them as PDF, and close them

[![Direct](https://img.shields.io/badge/Direct%20Link-OpenSaveAsPDFAndClose.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/OpenSaveAsPDFAndClose.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OpenSaveAsPDFAndClose.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Opens the chosen files in Illustrator one by one, saves each as PDF with the PDF preset, bleed and trim marks chosen in a dialog, and closes it.

Based on the sample script "Save as PDFs.jsx" bundled with Illustrator. The bundled version works on open documents; this one works on files that are not open.

### Main Features

- Choose several target files (.ai / .eps / .pdf / .svg) at once
- Choose a PDF preset
- Add bleed (mm, the same on all four sides) and trim marks (Japanese / Roman)
- Save next to the source file or into a chosen folder
- When a PDF of the same name exists: overwrite, add a number, or skip
- Runs through without stopping at font or link alerts
- Remembers the last settings
- Japanese and English UI

### How to Use

1. Run `OpenSaveAsPDFAndClose.jsx`.
2. Click "Choose..." in the Files panel and pick the target files.
3. Set the PDF options and the destination.
4. Click "Save". The files are opened, saved as PDF and closed in turn. At the end the number of saved files and any skipped or failed files are shown.

The PDF is named after the source file with its extension replaced by `.pdf` (e.g. `sample.ai` → `sample.pdf`).

### Files Panel

| Item | Description |
| --- | --- |
| Files | The number of chosen files. Hover to see the file names. |
| Choose... | Picks the target files. Choosing again replaces the previous selection. |

### PDF Panel

| Item | Description |
| --- | --- |
| PDF Preset | The PDF preset used for saving. Defaults to [Illustrator Default]. |
| Bleed | Adds the same bleed on all four sides (in mm). The document's bleed setting is not used. |
| Trim Marks | Adds trim marks, Japanese or Roman. |

### Destination Panel

| Item | Description |
| --- | --- |
| Same as Source File | Saves into the folder of the source file. |
| Custom | Saves into the folder chosen with "Choose...". The folder is created when missing. |
| Overwrite Existing PDF | Overwrites a PDF of the same name. |
| If Not Overwriting | "Add a Number" saves as the first free name, "name-1.pdf", "name-2.pdf"... "Skip" does not save. |

### Settings

The defaults can be changed in the "User Settings" block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `TARGET_FILES` | `[]` | Paths of the files targeted at first. Empty = choose with "Choose..." |
| `OPENABLE_FILE_PATTERN` | `/\.(ai\|eps\|pdf\|svg)$/i` | Files selectable with "Choose..." |
| `DEFAULT_PDF_PRESETS` | `["[Illustrator Default]", "[Illustrator 初期設定]"]` | Candidates for the initial PDF preset (the first one found is used) |
| `DEFAULT_SETTINGS` | — | Initial dialog values (later runs use the last ones) |

### Notes

- The source PDF itself and PDFs saved earlier in the same run are never overwritten, even with "Overwrite Existing PDF" on; "If Not Overwriting" applies instead.
- Files that are already open are skipped, because saving would switch the open document over to the PDF.
- With bleed or trim marks turned off, the PDF is saved without them rather than with the preset's settings.
- Opening a PDF opens only its first page, so for a multi-page PDF the saved PDF contains only page 1.
- Settings are stored in `illustrator-scripts/OpenSaveAsPDFAndClose.json` under `Folder.userData`.

### Update History

- v1.1.1 (2026-10-08) The source PDF itself and PDFs saved earlier in the same run are never overwritten (with "Overwrite Existing PDF" on, a number is added or the file is skipped)
- v1.1.0 (2026-10-08) Initial version
