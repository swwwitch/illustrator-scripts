# Pack objects onto artboards

[![Direct](https://img.shields.io/badge/Direct%20Link-PackItemsOnArtboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/PackItemsOnArtboards.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PackItemsOnArtboards.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Packs the selected items tightly onto the artboards. Items of different sizes are packed over and over with different sort orders, and the best result is kept. Spacing, margin, rotation and the number of attempts are set in a dialog.

### Main Features

- Fills the artboards in order. Items that do not fit on the first artboard move on to the next.
- **Space between items** and **Artboard margin** can be set (1 mm and 5 mm by default). Both accept negative values; a negative **Space between items** makes items overlap.
- **Max attempts** sets the maximum number of attempts. The dialog opens with a value calculated from the number of items.
- **Maximize** chooses which value counts most when comparing attempts: **Items packed** or **Area packed**.
- Each attempt sorts the items differently: by area, largest dimension, width, height, then alternating large and small. After that the order is random.
- **Use all attempts** keeps looking for a better result up to the maximum, even after every item fits.
- **Keep stacking order** packs once in stacking order, without sorting by size.
- **Allow 90° rotation** rotates an item by 90° when it does not fit otherwise.
- **Allow any rotation** rotates each item to its smallest bounding rectangle before packing.
- **Show results** shows the attempt used and the count per artboard afterwards (off by default).
- **Try Random** packs once in random order without closing the dialog.
- Items that could not be packed stay selected.

### Usage

1. Select the items.
2. Run the script.
3. Set the spacing, attempts and rotation.
4. Click **Pack**.

### Notes

- A group moves as a single item.
- Sizes are measured from the visible bounds, strokes included. Point text uses its outlined shape, and clipping groups use the area inside the mask.
- Items are packed by their bounding rectangles, so they do not nest into each other's shapes.
- After **Try Random**, the result stays even if you click **Cancel**.
- Many attempts or many items can take a while.

### Credits

Modified from [Bin Packing For Illustrator and Indesign](https://github.com/mark1bean/bin-packing-for-illustrator-and-indesign) by m1b (MIT License). The packing code is based on code that trentium published on [Stack Overflow](https://stackoverflow.com/questions/56642111/bin-packing-js-implementation-using-box-rotation-for-best-fit).

### Update History

- v1.2.0 (2026-11-07) : First public release. **Show results** is now off by default. The library is bundled, so this script runs on its own
