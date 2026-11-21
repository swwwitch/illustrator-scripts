# Convert the whole document to grayscale and delete swatches

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertDocumentToGrayscale.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ConvertDocumentToGrayscale.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertDocumentToGrayscale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Selects every object in the document and turns it monochrome with Edit > Edit Colors > Convert to Grayscale.
- Then deletes all swatches and removes swatch groups left empty.

### Usage

1. Open the target document
2. Run the script

### Notes

- There is no dialog. The whole document is processed regardless of the selection.
- Locked or hidden layers and objects are unlocked and shown while converting, then restored.
- Linked images, symbol definitions, the contents of patterns and brushes, and graphic styles are not converted.
- Swatches that cannot be deleted, such as [None] and [Registration], remain.
- To revert, run Undo several times.

### Change Log

- v1.0.0 (20261121): Initial release
