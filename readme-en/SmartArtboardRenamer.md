# SmartArtboardRenamer.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartArtboardRenamer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/SmartArtboardRenamer.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Batch renames artboards by combining a prefix and a suffix with reference text (the frontmost text frame, text on a chosen layer, the original artboard name, or a custom string).
- Every change to the settings refreshes the list on the right as an "Original Name -> New Name" preview.
- Artboards can also be reordered and renamed individually from the same dialog.

### Main Features

#### Rename conditions (left column)

- Token buttons insert elements into the prefix and suffix fields
    - Sequence (1 / 01), file name (#FN), date (#DT, yyyyMMdd), separators (- / _)
    - The "x" button clears the target field
- The number you type is the starting value and increases by 1 per artboard (zero padding such as "01" keeps its digit count)
    - The **first number** found in the prefix or suffix is treated as the sequence, so a fixed string like "A4_" also increments ("A4_" → "A5_" → "A6_")
- Four text source modes
    - **Original Artboard Name**: uses the name as it was when the dialog opened (not the post-Refresh name)
    - **Custom**: uses the entered string (leave it empty for no reference text)
    - **Frontmost Text**: recursively scans layers and groups and uses the first text frame found on each artboard
    - **Layer**: combines text from the selected layer hierarchy (including sublayers and groups)
- Hidden and locked layers and objects are ignored
- Appends "_1", "_2", etc. only when names collide (non-duplicates stay as-is), including collisions with artboards outside the target range
- When artboards overlap, a text frame belongs to the **smallest** artboard containing its center

#### Target artboards (right column, top)

- Choose "All" or "Range" (e.g. "1-3,5"); numbers outside the document are ignored and a reversed range like "5-1" is read as ascending
- Two-way linked with the Reorder / Rename checkboxes (toggling checkboxes updates the range text, and switches to "All" when every row is on)

#### Reorder / Rename (right column, bottom)

- Lists each artboard as "Original Name -> New Name", reflecting the live preview
- Checked rows allow a manual override of the new name (manual edits are never overwritten by the preview)
- Move checked rows with Top / Up / Down / Bottom (rename targets are rebuilt from the reordered positions)
- Option-click a checkbox to toggle all rows on or off
- A scrollbar appears when the document has more than 12 artboards (reordering scrolls to follow the moved rows)

#### Buttons

- **Refresh**: commits the current settings, manual edits, and reordering to the artboards immediately
- **Cancel**: restores the artboard names and order from before the dialog opened

### Process Flow

1. Set the prefix, suffix, text source, and target artboards in the dialog
2. Check the preview in the right column, then reorder or rename rows as needed
3. Use Refresh to commit intermediate results, and OK to apply the final settings

### Not Covered

- When no document is open (the script exits without doing anything)
- Hidden and locked layers and objects (excluded from text collection)
- When the specified layer is missing or hidden (an alert is shown and the rename is cancelled)

### Update History

- v1.0.0 (20250509): Initial version
- v1.0.1 (20250512): Improved layer reference and UI adjustments
- v1.5.2 (20260508): Refresh now commits right-column manual name edits and reorder changes
- v1.5.3 (20260508): On name collisions, the first occurrence now starts at "_1" (non-duplicates remain untouched)
- v1.5.4 (20260807): Unified the preview and rename logic, and cleaned up naming, layout, and comments
    - Fixed reordering being applied twice on Refresh then OK, which reverted the names of non-target artboards
    - Fixed the right column showing mismatched original names and previews after reordering and refreshing
    - Fixed checked rows targeting the wrong artboards after a reorder
    - Sequence numbers now follow the displayed row order instead of the canvas order when rows are reordered
    - Fixed the range text keeping a stale selection after every checkbox was cleared
    - Cancel now restores the artboard order (rects) as well as the names
    - Added a scrollbar so long artboard lists are no longer clipped
    - Fixed Layer mode joining text inside groups and sublayers twice
    - Fixed duplicate-name avoidance ("_1", "_2") being disabled whenever the prefix or suffix contained a number
    - Fixed Frontmost Text mode picking up another artboard's text when artboards overlap
    - Range input now handles out-of-range, duplicate, and reversed entries safely
