# Reorder artboards by name or position

[![Direct](https://img.shields.io/badge/Direct%20Link-ReorderArtboardsByPosition.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ReorderArtboardsByPosition.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReorderArtboardsByPosition.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Script to reorder artboards by name or position (Top Left, Top Right, Left Top, Right Top)
- Reorders the Artboards panel order based on the visual (canvas) arrangement on the document

![](https://www.dtp-transit.jp/images/ss-604-650-72-20250707-032528.png)

### Main Features

- Sort artboards by name or various position orders
- Adjust tolerance to allow slight misalignments
- Choose sorting method and tolerance via UI

### Process Flow

1. Select sort method and tolerance in dialog
2. Execute sorting with OK button
3. Changes are applied immediately

### Original / Credit

- m1b: https://community.adobe.com/t5/illustrator-discussions/randomly-order-artboards/m-p/12692397
- https://community.adobe.com/t5/illustrator-discussions/illustrator-script-to-renumber-reorder-the-artboards-with-there-position/m-p/12752568

### note

- https://note.com/dtp_tranist/n/nb416cb01728a

### Changelog

- v1.0.0 (20231115): Initial version (UI improvements and limit extension by Andrew_BJ)
- v1.1.0 (20231116): Added tolerance auto-calculation feature, slider support, and logic cleanup
- v1.2.0 (20260415): Refined the UI and structure as a dedicated Top Left reorder tool, and adjusted rearrange settings and preview display
- v1.3.0 (20260508): Added Artboard Names panel (auto row-column naming, reformat existing names, separator and digit width selection), split rearrange logic and UI construction into responsibility-based helper functions, unified localization behind a single helper, and refined linked column/row gap behavior
- v1.3.1 (20260508): Swapped the execution order of rearrange and panel reorder so the panel order matches the post-rearrange visual layout
- v1.4.0 (20260513): Replaced the panel order option with radio buttons (By name / Match canvas order / Keep as is); by-name uses a natural sort that zero-pads digit runs to 10 characters
- (20260807): Unified the overview and basic-info blocks with the shared format, reorganized layout constants and UI helpers, and added JSDoc to every function (no functional change)
- v1.5.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.5.1 (20260928): Replaced the Link checkbox with a link icon
- v1.5.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.5.2 (20260929): Dialog opacity changed to 98%
- v1.5.3 (20260929): Fixed the default tolerance sometimes merging rows into one. Locked or hidden objects and layers now move along when rearranging. Typed column/row gaps are clamped to the minimum. A failed rearrange now restores the original layout. Faster by-name sorting. Rearranging now keeps the artboards centered where they were (by column count used to push them to the canvas top-left; coordinates are now handled in document space, and the layout is kept on the canvas). Names that state the artboard size, such as 1920x1080, are no longer read as row-column, and an x right after a letter (as in "index2") is no longer a separator. With rearranging on, the preview shows the order after the rearrange
- v1.5.4 (20260930): Fixed an error when running with characters selected by the Type tool
