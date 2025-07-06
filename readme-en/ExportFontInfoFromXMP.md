# ExportFontInfoFromXMP.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-ExportFontInfoFromXMP.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/ExportFontInfoFromXMP.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script that extracts font usage information from XMP metadata embedded in an Illustrator document and exports it as text, CSV, or Markdown.
- You can select the output format via a dialog, or export all formats at once.

### Main Features

- Supports three formats: TXT, CSV, and Markdown
- CSV is output in UTF-16 with BOM
- Markdown escapes underscore (_) only
- Automatically renames if duplicate filenames exist
- Japanese and English UI support

### Process Flow

1. Extract font information from document XMP metadata
2. Select export format in the dialog
3. Save the font information on the desktop in the specified format

### Update History

- v1.0.0 (20250511): Initial version