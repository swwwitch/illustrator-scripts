# SlideCollage-mihiraki

[![Direct](https://img.shields.io/badge/Direct%20Link-SlideCollage--mihiraki.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SlideCollage-mihiraki.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SlideCollage-mihiraki.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Places a PDF/AI file over a given page range and lays each page out on its own artboard.

Landscape pages are treated as spreads, split left and right, and ordered according to the binding direction.

### Features

- Page range selection
- Automatic spread splitting, with left- and right-binding order
- PDF crop box selection
- The page count is estimated from the selected placed image or the chosen file and applied to the page range

### Usage

1. Run the script.
2. Choose the PDF/AI file and the page range.
3. Set the binding direction and the crop box, then run it.

### Notes

- The dialog position is remembered only for the session.
- Use PDFAISpreadImporter.jsx to expand into a new document instead.

### Update History

- v1.0 (2026-03-17)
