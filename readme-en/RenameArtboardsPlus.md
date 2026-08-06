# RenameArtboardsPlus.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-RenameArtboardsPlus.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/RenameArtboardsPlus.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Artboard names decide your exported file names, so they are not something to fix up later. Yet renaming them one by one in the Artboards panel stops being practical as soon as the document has more than a handful.

This script lets you assemble a naming rule from **the document name, a prefix, the original artboard name, sequential numbering, and a suffix**, check the result in a preview, and apply it to every artboard at once.

![](https://www.dtp-transit.jp/images/ss-1210-1120-72-20250713-080350.png)

## How to use

1. Open a document that has artboards
2. Run the script
3. Assemble the naming rule in each panel, check the preview, and click OK

If no document is open, the script exits without doing anything.

## How a name is assembled

The new artboard name is concatenated in this order:

```
[Prefix] + [Artboard Name & Number] + [Separator] + [Number] + [Suffix string]
```

When the numbering format is "None", the separator and the number are omitted.

## Prefix

| Item | Description |
| --- | --- |
| File Name | "Use" puts the document name (without its extension) at the front |
| Separator | Character inserted between the file name and what follows (None / - / _). Editable only while "Use" is selected |
| String | Any text |

## Artboard Name & Number

| Choice | Resulting text |
| --- | --- |
| None | (nothing) |
| Number | The artboard's position (1, 2, 3, ...) |
| Name | The artboard name before renaming |
| Number-Name | `1-original name` |
| Number_Name | `1_original name` |

The "Number" here is the artboard's own position. It is not affected by the Start Number or Increment in the Suffix panel.

## Suffix

| Item | Description |
| --- | --- |
| Separator | Character placed right before the number (None / - / _). Dimmed when the numbering format is "None" |
| Numbering Format | None / Number / Alphabet (Upper) / Alphabet (Lower) |
| Start Number | Where the numbering starts — `1` or `001` for numbers, `A` or `ab` for letters |
| Increment | Step between numbers. Only active when the numbering format is "Number" |
| String | Any text appended at the end |

### Zero padding

The number of digits you type into Start Number becomes the padding width. Enter `001` and you get `001`, `002`, `003`, ...; enter `1` and you get `1`, `2`, `3`, ... There is no separate field for the digit count.

### Alphabetical numbering

Letters run `A` → `B` → ... → `Z` → `AA` → `AB` → ... The increment is not used; letters always advance by one. The start value can be a single letter such as `A` or several letters such as `ab`.

## Preview

Every change refreshes the preview, showing up to 15 results from the top. Anything beyond that is collapsed into "... (more)".

If the input is invalid — for example, non-digit characters in Start Number or Increment — the preview shows "※ Invalid number" and OK reports the error instead of proceeding. Surrounding whitespace is ignored.

## Presets

Choosing a built-in preset from the dropdown fills in every field at once.

| Preset | Example result |
| --- | --- |
| ファイル名+連番3 | `filename-001` |
| アートボード名と連番 | `original name-1` |

Export Preset saves the current settings as a single line of text (UTF-8). The file name you choose becomes the preset name.

Paste the exported text into `BUILTIN_NAMING_PRESETS` in the script to add it as a built-in preset. There is no import-from-file feature.

## Notes

- If a name would come out empty while numbering is off, the script warns and does not rename
- Names are always built from the artboard names as they were before renaming
- If Illustrator rejects a name, the script warns and leaves the dialog open
- Japanese and English UI, switched automatically by the Illustrator locale
- Compatible with ExtendScript (ES3)

## Update History

- v1.3.1 (2026-08-06): Write exported presets as UTF-8 so Japanese labels are not garbled, dim the suffix separator when the numbering format is "none" (it has no effect there), reject non-digit input and ignore surrounding whitespace in start number / increment, keep the dialog open when renaming fails, unified panel margins and spacing through a shared layout setup (setupPanel / PANEL_MARGINS), renamed variables / functions / panels to match what they actually represent, grouped LABELS into categories, and added JSDoc to every function
- v1.3.0 (2026-05-09): Added a warning when the resulting name would be empty, tagged separator radio buttons with their values to remove array-order dependency, escaped strings when exporting presets, changed the default start number to "001"
- v1.2.0 (2026-05-07): Localization overhaul (L / labelText / labelWithCount), key-based internals for EN locale support, added "none" numbering format, removed the Apply button, split main() into UI builders / preset I/O / pure helpers
- v1.1 (2025-04-30): Added auto-detection of padding digits, simplified preset labels, enhanced ES3 support
- v1.0 (2025-04-20): Initial version created
