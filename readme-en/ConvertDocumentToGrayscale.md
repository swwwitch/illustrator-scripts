# Turn the whole document into CMYK monochrome (text in K100)

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertDocumentToGrayscale.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ConvertDocumentToGrayscale.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertDocumentToGrayscale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Switches the document to CMYK, then converts every object to grayscale. Useful for preparing artwork for monochrome printing.
- Text becomes K100, and the swatches are deleted at the end.

### Features

The script runs these steps in order.

1. Converts object colors to CMYK with Edit > Edit Colors > Convert to CMYK
2. Switches an RGB document to File > Document Color Mode > CMYK Color
3. Converts the contents of symbols to grayscale (nested symbols too)
4. Turns every object monochrome with Edit > Edit Colors > Convert to Grayscale
5. Sets the fill and stroke of text to K100 (text inside symbols too)
6. Deletes all swatches and removes swatch groups left empty

Locked or hidden layers and objects are unlocked and shown while processing, then restored.

### Usage

1. Open the target document
2. Run the script

### Notes

- There is no dialog. The whole document is processed regardless of the selection.
- Only solid fill and stroke colors of text become K100. White (0% gray) stays white (set KEEP_WHITE_TEXT at the top of the script to false to make it K100 too). Text with a gradient or pattern stays grayscale.
- Symbol instances are swapped to new symbols whose contents are converted to grayscale, and the new symbols replace the originals. Symbol options such as the registration point, the dynamic type and 9-slice scaling are not carried over.
- Linked images, the contents of patterns and brushes, and graphic styles are not converted.
- Swatches that cannot be deleted, such as [None] and [Registration], remain.
- To revert, run Undo several times.

### Change Log

- v1.0.0 (20261121): Initial release
- v1.1.0 (20261121): Added Convert to CMYK and the switch to CMYK document color mode, grayscale conversion of symbol contents, and K100 text
