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
- Text is measured from an outlined copy, so the export area follows the actual glyphs (the original text is left untouched)
- Background: transparent, black, white, transparency grid or a color code. The grid tile size is given as a percentage, where 100% is an 8pt square
- Color codes accept three formats: `#RRGGBB`, `R255G255B255` and `C0M100Y100K0`
- Margin: top, bottom, left and right set individually, in the current ruler unit. Same Value applies the top value to all four sides
- Rounding picks how the export area is rounded: Optimize to pixel grid, Round values in current unit, or Do nothing
- Border: enabled by the Width checkbox, with a width plus color (black, white or a color code), drawn inside the export area with a minimum of 1px
- Export size: 1x to 4x, a custom scale (%) or a target width (px). The 1x–4x labels show the resulting pixel size including the margin, and follow the margin as it changes
- Filename built from the document name (used or ignored), a delimiter (none, `-`, `_`) and a suffix, with a live preview (the delimiter is dropped when the document name is ignored)
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
| Margin | Top / Bottom / Left / Right, in the current ruler unit; Linked keeps all four equal; Rounding picks how the area is rounded |
| Border | Width checkbox plus a width in the current ruler unit, and black, white or a color code |
| Export Size (px) | 1x–4x / custom scale (%) / target width (px) |
| Export Filename | Whether to use the document name, the delimiter, and the Suffix checkbox with its value |
| Export Location | Desktop / same folder as the file, and whether to show the folder afterwards |

### Notes

- The export area is grown outward according to Rounding, so the artwork is never clipped. The default, Optimize to pixel grid, rounds to whole points, and at 100% one pt equals one px.
- The border is drawn inside the export area, so add a margin as well when you need clearance around the artwork.
- For an unsaved document, "Same as File" falls back to the desktop.
- "Show Folder After Export" is macOS only.
- Preset margins and border widths are kept in mm; when the ruler unit is something else, the converted value is filled in. Save Preset converts back to mm.
- A background or border color code that cannot be read is simply not drawn: the background stays transparent and the border is omitted.
- The export scale is capped at 776.19%. When a target width asks for more, the image is exported at the cap and the applied scale is reported.
- A very small transparency grid percentage would produce a huge number of tiles, so the tile size is enlarged automatically past a certain point.
- Text is outlined on a temporary copy only for measuring, and the copy is removed afterwards, so the original text is never modified.
- Layers are hidden temporarily while the script runs, so the canvas changes during the process and is restored afterwards.

### Article

[【Illustrator】選択したオブジェクトを書き出すスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/necf308c39f5d)

### Update History

- v1.0.2 (2026-09-11): Fixed stacking order across layers, the working layer left behind when something failed mid-run, deletion of same-named user artwork, the crash on a text-tool selection, the destination folder for an unsaved document, the delimiter with no document name, how an unreadable color code is handled for the background and border, and the transparency grid spilling past the export area. Preset margins and border widths are now kept in mm, the UI wording was revised, and tooltips were added throughout
- v1.0.1 (2026-09-11): Margins are now set per side with a Linked option, Rounding picks how the export area is rounded, the border and the suffix are toggled with checkboxes, text is measured from an outlined copy, the transparency grid tile no longer depends on the ruler unit (an 8pt square at 100%), and the export size labels are no longer clipped
- v1.0.0 (2026-09-11): Reorganized internals (shared settings reader, preview and export). Fixed the scale calculation for a target width, the error when a preset was selected, the RGB/CMYK color code formats, the custom scale radio that could not be selected, the margin in the scale labels, and the appearance on a dark UI
- v0.5.0 (2025-06-19): Initial version
