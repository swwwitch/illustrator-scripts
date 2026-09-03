# Align vertically while toggling glyph bounds

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartVerticalAlign.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartVerticalAlign.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartVerticalAlign.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Aligns objects vertically — top, center or bottom — while toggling Align to Glyph Bounds for point text and area text. Every checkbox and radio button refreshes the preview, so you can compare the result with and without glyph bounds before committing.

### Main Features

- Toggle Align to Glyph Bounds for point type and area type independently
- Choose the alignment position (top, center, bottom)
- A Preview Bounds toggle decides whether strokes and effects count toward the bounds
- Every checkbox and radio button refreshes the preview immediately
- Hotkeys (T / M / B) switch the alignment position
- Japanese / English UI

### Usage

1. Select the objects you want to align.
2. Run the script.
3. Set Align to Glyph Bounds and the alignment position, check the preview, and click OK.

### Options

**Align to Glyph Bounds**

| Item | Description |
| --- | --- |
| Point Text | Align to Glyph Bounds for point type (preference `EnableActualPointTextSpaceAlign`) |
| Area Text | Align to Glyph Bounds for area type (preference `EnableActualAreaTextSpaceAlign`) |

Area Text is dimmed when the selection is point type, and Point Text is dimmed when it is area type.

**Alignment**

| Item | Description |
| --- | --- |
| Top | Align to the top |
| Center | Align to the center |
| Bottom | Align to the bottom |

**Preview Bounds**

| Item | Description |
| --- | --- |
| Preview Bounds | When ON, strokes and effects count toward the bounds (preference `includeStrokeInBounds`). Default OFF |

**Hotkeys**

| Key | Description |
| --- | --- |
| T / M / B | Top / Center / Bottom |

### Notes

- Text frames, paths, groups and compound paths are treated as alignable objects.
- The defaults depend on the selection: Bottom for text only, Center when text and shapes are mixed or when only shapes are selected.
- The checkboxes write Illustrator's preferences directly. Cancel does not restore their previous state.
- Alignment runs through Illustrator's menu commands, so every action — keyboard shortcuts included — is recorded in the undo history.

### Article

https://note.com/dtp_tranist/n/n9ee716675032

### Update History

- v1.1.1 (2026-09-03) Fixed the preview not updating when Align to Glyph Bounds is turned back OFF; moved alignment positions, shortcuts and preference keys into tables; cleaned up naming and structure
- v1.1 (2025-08-04) Adjusted the logic used when the dialog opens
- Initial release (2025-08-04)

### Script info

- Version: v1.1.1
- First release: 2025-08-04
- Last updated: 2026-09-03
