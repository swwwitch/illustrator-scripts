# Export the selected objects as PNG

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartObjectExporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/SmartObjectExporter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectExporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

![](https://www.dtp-transit.jp/images/ss-1244-1296-72-20250713-080243.png)

### Overview

Places the selected objects on a temporary artboard and exports them as PNG.

The background, margin, border, export size and filename are set in a dialog, and every setting is previewed on the artboard itself. Frequently used combinations can be recalled as presets.

### Main Features

- Copies the selection onto a working layer, so nothing else appears in the exported image
- Background: transparent, black, white, transparency grid or a color code. The grid tile size is given as a percentage, where 100% is 10 ruler units
- Color codes accept three formats: `#RRGGBB`, `R255G255B255` and `C0M100Y100K0`
- Margin: none, horizontal, vertical or all sides, in the current ruler unit
- Border: width plus color (black, white or a color code), drawn inside the export area with a minimum of 1px
- Export size: 1x to 4x, a custom scale (%) or a target width (px). The 1x–4x labels show the resulting pixel size including the margin, and follow the margin as it changes
- Filename built from the document name (used or ignored), a delimiter (none, `-`, `_`) and a suffix, with a live preview
- The chosen scale or target width is reused as the suffix automatically
- Characters not allowed in filenames (`¥ / : * ? " < > |`) and whitespace are replaced with the delimiter
- Export location: desktop or the folder of the current document
- Opens the destination folder after the export (macOS only)
- Four built-in presets, plus a command that writes the current settings to a text file
- Numeric fields step with the arrow keys (shift for 10, option for 0.1)

### How to Use

1. Select the objects you want to export.
2. Run the script.
3. Set the background, margin, border, export size, filename and destination in the dialog. The settings are previewed on the artboard.
4. Click OK to export. Cancel discards the preview and leaves the document untouched.

### Options

| Item | Description |
| --- | --- |
| Preset | Recalls a built-in setting. Save Preset writes the current settings to the desktop as text |
| Background | Transparent / Black / White / Transparency grid (%) / Color code |
| Margin | None / Horizontal / Vertical / All sides, in the current ruler unit |
| Border | None / Add (width plus black, white or a color code) |
| Export Size (px) | 1x–4x / custom scale (%) / target width (px) |
| Export Filename | Whether to use the document name, the delimiter and the suffix |
| Export Location | Desktop / same folder as the file, and whether to show the folder afterwards |

### Notes

- The export area is rounded to whole units so the result is pixel perfect. At 100% one pt equals one px.
- The border is drawn inside the export area, so add a margin as well when you need clearance around the artwork.
- For an unsaved document, "Same as File" falls back to the desktop.
- "Show Folder After Export" is macOS only.
- The export scale is capped at 776.19%. When a target width asks for more, the image is exported at the cap and the applied scale is reported.
- A very small transparency grid percentage would produce a huge number of tiles, so the tile size is enlarged automatically past a certain point.
- Layers are hidden temporarily while the script runs, so the canvas changes during the process and is restored afterwards.

### Article

[【Illustrator】選択したオブジェクトを書き出すスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/necf308c39f5d)

### Update History

- v1.0.0 (2026-09-11): Reorganized internals (shared settings reader, preview and export). Fixed the scale calculation for a target width, the error when a preset was selected, the RGB/CMYK color code formats, the custom scale radio that could not be selected, the margin in the scale labels, and the appearance on a dark UI
- v0.5.0 (2025-06-19): Initial version
