# AddPageNumberFromTextSelection.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-AddPageNumberFromTextSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddPageNumberFromTextSelection.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that duplicates and places page-number text on every artboard, using the selected point text as a template.
- Start number, prefix, suffix, zero padding, and total page display are configurable, and every change is reflected in the live preview immediately.
- The text is placed on the `_pagenumber` layer, which is created automatically when it does not exist.

<img alt="" src="https://www.dtp-transit.jp/images/ss-672-346-72-20250713-081024.png" width="50%" />

### Main Features

- Generates page numbers that inherit the styling of the selected point text
- Start number, prefix, and suffix
- Zero padding, matched to the digit count of the total (`1` → `01`)
- Total page display (`3/12` format)
- Live preview, with restoration on Cancel
- Auto-create, temporary unlock, and state restoration for the `_pagenumber` layer
- Tooltips on every control
- Japanese and English UI support

### Process Flow

1. Select a single point text object to use as the page-number template (any layer will do)
2. Configure the start number, prefix, suffix, and other options in the dialog
3. Check the preview
4. Click OK to place page numbers on all artboards (Cancel discards the preview)

### Dialog

| Item | Description |
| --- | --- |
| Prefix | Text placed before the number (e.g. `P.`) |
| Start number | Number assigned to the first artboard |
| Zero padding | Pads with `0` to match the digit count of the total (e.g. `1` → `01`) |
| Suffix | Text placed after the number (e.g. `page`) |
| Show total pages | Shows the number as "current/total" (e.g. `3/12`) |

### Keyboard Shortcuts

| Key | Action |
| --- | --- |
| Up / Down | Steps the start number by 1 |
| Shift + Up / Down | Snaps the start number to a multiple of 10 |
| Z | Toggles zero padding |
| A | Toggles the total page display |

Z and A do not toggle while a text field has focus — typing takes precedence there.

### Notes

- Point text only. The script does not change paragraph alignment or styling.
- Edits to the prefix, suffix, and start number reach the preview once the field is committed, by pressing Tab or moving focus to another control.
- During preview and on commit, existing text on the `_pagenumber` layer is replaced by copies of the template.
- If the `_pagenumber` layer did not exist beforehand, it is removed on Cancel. On OK it is kept.

### Update History

- v2.1.0 (2026-07-27): Unified UI wording around "page number". Added tooltips. Checkbox toggles and committed field edits now refresh the preview. Fixed detection of a nested `_pagenumber` layer, excluded off-artboard text from numbering, and made a failed cut abort instead of deleting the existing text. Consolidated internal routines.
- v2.0.1 (2026-05-16): Internal cleanup. Improved template detection on commit, preview undo tracking, and `_pagenumber` restoration.
- v2.0 (2026-01-08): Added rollback-based preview management.
- v1.9 (2025-08-10): Added auto-create, temporary unlock, and restoration for the `_pagenumber` layer.
- v1.8 (2025-06-25): Added live preview and all-artboards duplication.
- v1.0 (2025-06-25): Initial release
