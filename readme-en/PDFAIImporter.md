# PDFAIImporter

[![Direct](https://img.shields.io/badge/Direct%20Link-PDFAIImporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/PDFAIImporter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PDFAIImporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Imports a PDF or AI file over a given page range and places the pages on the current document.
Placement is either "one artboard per page" or "as objects, without adding artboards".

### Target

- All pages (default)
- First page only
- Specified pages (for example `1-10` or `1,3,5`)

The target-page and placement panels stay disabled until the source file is settled.
The total page count is estimated from the selected placed image, or from the PDF/AI chosen with the file button, and the target-page panel shows "pages actually placed / total pages".
The file dialog offers PDF and AI files only.

### Placement

- Per artboard: creates or updates an artboard to match each page size and places it at a fixed 100% scale
- Ignore artboards: places the pages as objects at the given scale, anchored to the top-left of the current artboard, then selects them and fits the view

### Article

https://note.com/dtp_tranist/n/n42595650216f

### Script info

- Version: v1.1.1
- Last updated: 2026-04-13
