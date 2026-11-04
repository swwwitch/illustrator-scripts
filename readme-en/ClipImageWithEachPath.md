# Cut an image into pieces with paths

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipImageWithEachPath.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipImageWithEachPath.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipImageWithEachPath.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Duplicates the selected image once for each selected path and clips each copy with its own path. Use it to cut several illustrations drawn on one image into separate objects. When the selection includes lines (open paths), the image area is divided by the paths and each resulting region becomes a mask.

### Main Features

- Duplicates the image once per path and makes a clipping mask from that path.
- Each clip group takes the stacking position of the path used as its mask.
- Paths such as rectangles and ellipses, and compound paths, can be used as masks.
- Works with both linked and embedded images.
- Draw lines (open paths) across the image to clip each region the lines divide it into (divided the same way as Pathfinder Divide).
- Removes the original image at the end and leaves the new clip groups selected.

### Usage

1. Draw paths over the image around each area you want to cut out (as many as you like).
2. Select the image and the paths together.
3. Run the script. No dialog box appears.

To divide with lines, draw lines that cross the image from edge to edge instead of the enclosing paths in step 1. Let both ends of each line run past the image.

### Notes

- The selection may contain only one image.
- Paths that do not overlap the image would make empty masks, so they are left untouched.
- The fill and stroke of each path disappear when it becomes a clipping mask.
- When the selection includes even one line, closed paths are used as dividers too (the inside and outside of a closed path become separate regions). Clip groups for divided regions take the stacking position of the original image.
- The lines and paths used as dividers are removed.
- Text, groups and guides in the selection are ignored. Paths inside groups are not used either.

### Update History

- v1.1.0 (2026-11-04): Lines (open paths) across the image now divide it into regions, each clipped separately
- v1.0.0 (2026-11-03): Initial version
