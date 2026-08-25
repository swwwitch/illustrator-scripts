# PDFAISpreadImporter

[![Direct](https://img.shields.io/badge/Direct%20Link-PDFAISpreadImporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/PDFAISpreadImporter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PDFAISpreadImporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Imports a PDF/AI file over a given page range and places each page on its own artboard in a new document.

Landscape pages are detected as spreads and split into two artboards, left and right.

### Features

- Page range selection (all pages / first page only / specific pages)
- Automatic spread detection for landscape pages, split left and right
- Even-page position selectable as right or left
- PDF crop box selection (Art / Crop / Trim / Bleed)
- Color mode of the new document selectable as CMYK or RGB

### Usage

1. Run the script.
2. Choose the PDF/AI file to import (a selected placed image is used if there is one).
3. Set the page range, the even-page position, the crop box and the color mode.
4. Run it, and the artboards are laid out in a new document.

### Notes

- The page count is estimated from the selected placed image or the chosen file and applied to the page range.
- The artboard spacing is fixed at 100 pt.
- The raster effects resolution is fixed at 300 ppi.
- Use PDFAIImporter.jsx when you do not want the pages split.

### Update History

- v1.1.0 (2026-03-18)
