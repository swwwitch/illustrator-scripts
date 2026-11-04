# Cut an image into pieces with paths

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipImageWithEachPath.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipImageWithEachPath.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipImageWithEachPath.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Duplicates the selected image once for each selected path and clips each copy with its own path. Use it to cut several illustrations drawn on one image into separate objects. When the selection includes lines across the image (open paths whose ends are far apart), the image area is divided by the paths and each resulting region becomes a mask.

### Main Features

- Duplicates the image once per path and makes a clipping mask from that path.
- Each clip group takes the stacking position of the path used as its mask.
- Paths such as rectangles and ellipses, and compound paths, can be used as masks. Unclosed paths drawn around an area, for example with the Pencil tool, work as enclosures too.
- Works with both linked and embedded images.
- Draw lines (open paths) across the image to clip each region the lines divide it into (divided the same way as Pathfinder Divide). The clip groups for divided regions are ordered left to right (top to bottom within a column).
- Paths and images inside selected groups are used too (not those inside clip groups).
- Removes the original image at the end and leaves the new clip groups selected.

### Usage

**Enclose with paths**

1. Draw paths over the image around each area you want to cut out (as many as you like).
2. Select the image and the paths together. The paths may stay grouped.
3. Run the script. No dialog box appears.

**Divide with lines**

1. Draw lines that cross the image from edge to edge. Slanted lines and lines with bends are fine. Letting both ends run past the image is safest, but lines that stop within 3 pt of the edge are extended to reach it.
2. Select the image and the lines together.
3. Run the script.

### Notes

- The selection may contain only one image.
- Paths that do not overlap the image would make empty masks, so they are left untouched.
- The fill and stroke of each path disappear when it becomes a clipping mask.
- Only open paths whose ends are far apart count as dividing lines (the distance between the ends is more than half the longer side of the path). An unclosed path drawn around an area, for example with the Pencil tool, has its ends almost on top of each other and is used as a mask.
- Dividing lines are extended by 3 pt at both ends before dividing. A line stopped against another line only pokes out slightly and adds no extra region.
- Tiny regions (under 1 pt²) and very thin ones (average width under 0.5 pt), which come from lines barely crossing each other, are not clipped.
- When the selection includes even one dividing line, closed paths are used as dividers too (the inside and outside of a closed path become separate regions). Clip groups for divided regions take the stacking position of the original image.
- The lines and paths used as dividers are removed, along with any groups left empty.
- Text and guides in the selection, and locked or hidden items inside groups, are ignored.

### Article

https://note.com/dtp_tranist/n/n6b28a104130d

### Update History

- v1.2.0 (2026-11-04): Fixed unclosed enclosing paths (for example drawn with the Pencil tool) being treated as dividing lines, which also turned the background into a region; only open paths whose ends are far apart now divide. Dividing lines are extended by 3 pt at both ends so lines stopping short of the edge still divide. Tiny or very thin regions are skipped. Paths and images inside groups are used. Clip groups for divided regions are ordered left to right (top to bottom within a column)
- v1.1.0 (2026-11-04): Lines (open paths) across the image now divide it into regions, each clipped separately
- v1.0.0 (2026-11-03): Initial version
