# GridArrangeArtboards

[![Direct](https://img.shields.io/badge/Direct%20Link-GridArrangeArtboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/GridArrangeArtboards.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that parses artboard names and rearranges the artboards as a row/column grid.
- Names that fully match "row number + separator (-, _, x) + column number" are treated as row/column targets (e.g. 1-1, 1_2, 2x1).
- Only when that pattern does not match, "prefix + separator + number" names are treated as prefix targets (e.g. banner-1, iconx3) and placed below the numeric grid, grouped into one row per prefix.
- Objects sitting on an artboard are moved by the same delta as the artboard.

### Main Features

- Spacing panel for column and row gaps; "Link" mirrors the column gap to the row gap
  - Input follows the document ruler unit (rulerType)
  - Default gap is the active artboard width divided by 8, rounded to an integer
  - Arrow keys step the value (Shift snaps to multiples of 10, Option steps by 0.1)
- Two exception modes for unmatched / duplicate names
  - **Append to current row**: unmatched names go to the end of the most recently matched row (from max column + 1 onward)
  - **Collect in final row**: unmatched names are laid out in the final row in the order they appear
  - The gap is tripled at the boundary into the exception area for clearer separation
- "Exclude Locked / Hidden" skips locked or hidden layers and/or objects when moving artwork
- "Reorder in Artboards panel" rebuilds the panel order in ascending row-then-column order
- Column widths and row heights use the per-column / per-row maximum size, so artboards of mixed sizes never overlap
- Runs "Fit All in Window" after processing
- Automatic Japanese / English UI

### Workflow

1. Scan the artboard names and classify them as row/column, prefix-number, or unmatched
2. Set gaps, exception mode, exclusions, and panel reordering in the dialog
3. Assign row/column numbers and compute cumulative offsets from the per-column max width and per-row max height
4. Translate the objects on each artboard first, then apply the new artboard rects (and rebuild the panel order if requested)

### Not Supported

- No open document
- Documents where no name matches the "row-column" or "prefix-number" pattern (an alert is shown and the script exits)
- Row 0, column 0, and number 0 (treated as invalid)
- Numeric-only prefixes (excluded to avoid ambiguity with row/column names)
- Individual items inside a group (groups move as a unit)
- Objects whose center falls outside the original artboard rect
- Locked or hidden layers and objects when the exclusion options are on

### Update History

- v1.1.0: Current version
