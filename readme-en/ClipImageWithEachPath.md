# Cut an image into pieces with paths

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipImageWithEachPath.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipImageWithEachPath.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipImageWithEachPath.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Duplicates the selected image once for each selected path and clips each copy with its own path. Use it to cut several illustrations drawn on one image into separate objects.

### Main Features

- Duplicates the image once per path and makes a clipping mask from that path.
- Each clip group takes the stacking position of the path used as its mask.
- Paths such as rectangles and ellipses, and compound paths, can be used as masks.
- Works with both linked and embedded images.
- Removes the original image at the end and leaves the new clip groups selected.

### Usage

1. Draw paths over the image around each area you want to cut out (as many as you like).
2. Select the image and the paths together.
3. Run the script. No dialog box appears.

### Notes

- The selection may contain only one image.
- Paths that do not overlap the image would make empty masks, so they are left untouched.
- The fill and stroke of each path disappear when it becomes a clipping mask.
- Text, groups and guides in the selection are ignored. Paths inside groups are not used either.

### Update History

- v1.0.0 (2026-11-03): Initial version
