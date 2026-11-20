# Swap objects (floating palette)

[![Direct](https://img.shields.io/badge/Direct%20Link-SwapNearestItemPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SwapNearestItemPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapNearestItemPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Swaps the selected object with the nearest object in a chosen direction (up, down, left or right).
A floating palette stays open and the direction is picked with the arrow keys or the palette's arrow buttons.

### Main Features

- While the palette is active, the up/down keys swap with the nearest object above or below, and the left/right keys with the nearest object to either side
- The palette's arrow buttons do the same (clicking one also activates the palette, so the arrow keys work right after)
- Pressing the same key again keeps moving the selected object one step further in that direction
- Widths (or heights) and the gap are kept, so the pair still occupies the same range
- Nearness is the gap along the search direction; only objects whose range overlaps on the perpendicular axis are candidates
- The selection can be changed on the artboard while the palette stays open
- When there is nothing to swap with, the palette shows "Nothing found" (no alert)
- Esc or Enter closes the palette. It reopens where it was last closed and moves sideways to avoid covering the selection

### Usage

1. Select one object and run the script
2. Click an arrow button on the palette, or click the palette to make it active and press an arrow key
3. After selecting another object, click an arrow button (while the palette is not active, the arrow keys nudge objects in Illustrator)

### Targets

- Can be swapped: paths, compound paths, groups, text, placed and embedded images, symbols, meshes, plugin items, graphs
- Excluded: guides, locked or hidden objects, objects on locked or hidden layers (including parent layers), objects inside groups or compound paths
- Clipping groups are measured by their masks (hidden parts do not count)

### Notes

- This is a persistent palette: after editing the script, close the palette before running it again
- The swap runs in the main engine through BridgeTalk (a persistent palette cannot touch the document directly), so it responds a little slower than the dialog version. The swap code is sent only on the first press after the palette opens; later presses send a short call

### Related scripts

- [Swap objects (dialog version)](SwapNearestItemWithDialogbox.md): the same swap in a modal dialog, which responds faster

### Update History

- v1.1.0 (2026-11-13) Created as a floating-palette version of the dialog version (SwapNearestItemWithDialogbox.jsx). The selection can be changed while the palette is open. Added arrow buttons. The version is shown at the bottom of the window. Added to the targets of CloseAllPalettes
- v1.0.8 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.0.7 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.0.6 (20260929): Dialog opacity changed to 98%
- v1.0.5 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%. Clip groups are measured by their masks (hidden parts no longer count when searching and swapping)
- v1.0.4 (20250615): Removed temporary grouping, cleaned up logic
- v1.0.3 (20250614): Added temporary group handling for multiple selection
- v1.0.2 (20250613): Refactored with getCenter() and getSize()
- v1.0.1 (20250612): Added exclusion of group/compound path children and lock support
- v1.0.0 (20250610): Initial release
