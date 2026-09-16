# Export the font information used in a document

[![Direct Link](https://img.shields.io/badge/Direct%20Link-ExportFontInfoFromXMP.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ExportFontInfoFromXMP.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExportFontInfoFromXMP.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

<img alt="" src="https://www.dtp-transit.jp/images/ss-476-456-72-20250713-081802.png" width="50%" />

## Overview

Extracts all font usage information from the XMP metadata embedded in the active Illustrator document and exports it as a tab-separated text file, CSV, or Markdown.

Composite fonts are reported together with their member fonts.

The script also runs on an unsaved document, or one edited since the last save. In that case fonts absent from the XMP are topped up by scanning the text in the document.

## Main Features

- Supports three formats: TXT, CSV, and Markdown (all three can be exported at once)
- Destination can be the desktop or the document's own folder
- Opens the destination folder after exporting (on by default)
- Runs on unsaved or edited documents, topping up the list from the text
- Can narrow the export down to missing fonts only
- Lists the member fonts of composite fonts
- CSV is written in UTF-16 with BOM
- Markdown escapes underscore (`_`) only
- Automatically renames with a serial number when a duplicate filename exists
- Radio buttons can be selected with the arrow keys
- Japanese and English UI

## How to Use

1. Open the document whose font info you want to export.
2. Run `ExportFontInfoFromXMP.jsx`.
3. Choose the export format, destination, and options in the dialog.
4. Click OK.

The output filename is the document name plus `_fontInfo`, with the extension of the chosen format (e.g. `sample_fontInfo.csv`).

## Export Format Panel

| Item | Description |
| --- | --- |
| Text File (.txt) | Writes a tab-separated text file. |
| CSV File (.csv) | Writes a UTF-16 CSV file with BOM, ready to open in Excel. |
| Markdown File (.md) | Writes a Markdown file with headings. |
| All Formats (TXT + CSV + MD) | Writes all three formats at once. |

## Destination Panel

| Item | Description |
| --- | --- |
| Desktop | Saves to the desktop. |
| Same folder as the file | Saves to the same folder as the document. Unavailable on an unsaved document. |
| Open the folder after exporting | Reveals the destination folder in Finder / Explorer after exporting. On by default. When off, the exported filenames are shown in an alert instead. |

## Options Panel

| Item | Description |
| --- | --- |
| Missing fonts only | Exports only the fonts that are not installed. When there are none, an alert is shown and nothing is written. Off by default. |

## Output Example

Markdown

```markdown
### sw-B

- fontName: ATC-73772d42
- fontFace: 
- fontType: Composite Font
- fileName: sw-B

#### Composite Fonts

- RyoGothicStd-Bold.otf
- NotoSansCJKjp-Light.otf
- HiraginoSans W2.ttc
- RyoGothicStd-Heavy.otf
```

Text file

```text
fontName:	ATC-73772d42
fontFamily:	sw-B
fontFace:	
fontType:	Composite Font
fileName:	sw-B
Composite Fonts:
- RyoGothicStd-Bold.otf
- NotoSansCJKjp-Light.otf
- HiraginoSans W2.ttc
- RyoGothicStd-Heavy.otf
```

## Settings

The defaults can be changed in the "User Settings" block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `FILENAME_SUFFIX` | `"_fontInfo"` | Suffix appended to the output filename |
| `SECTION_DIVIDER` | `"-----------------------------"` | Section divider used in the text file |
| `OPEN_FOLDER_DEFAULT` | `true` | Default for "Open the folder after exporting" |
| `MISSING_ONLY_DEFAULT` | `false` | Default for "Missing fonts only" |

## Notes

- Font info comes from the saved XMP. On an unsaved or edited document the XMP still holds the last-saved state, so the missing part is topped up from the text in the document.
- Entries topped up from the text have an empty `version` and `fileName`, and are not treated as composite fonts, so no member list is written for them. **Save the document first if you need that information.**
- The text scan runs character by character, so a large unsaved or edited document takes a while. A saved document is not scanned at all.
- If no document is open, or no font information is found, the script shows an alert and exits.
- No `version` is written for composite fonts.
- "Missing fonts only" decides by matching names against the installed fonts. A composite font is judged by its member fonts; when those cannot be read it cannot be judged and is left out.

## Article

[Exporting the fonts used in a document (Japanese) | DTP Transit](https://note.com/dtp_tranist/n/n16e7e95652b6)

## Update History

- v1.0.3 (2026-09-17): Added support for unsaved and edited documents (topping the list up from the text). Added "Missing fonts only". Centered the button row
- v1.0.2 (2026-08-06): Added "Open the folder after exporting". Fixed a dropped composite member font, arrow-key selection, XML entity decoding, and CSV escaping
- v1.0.1 (2026-06-17): Added destination choice (desktop / same folder), panel layout, and unsaved-document check
- v1.0.0 (2025-05-10): Initial version
