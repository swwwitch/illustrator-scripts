# NewGuideMaker


[![Direct](https://img.shields.io/badge/Direct%20Link-NewGuideMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/NewGuideMaker.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

- Script to create guides in Illustrator by specifying direction, position, unit, and target (canvas or artboard) via dialog

<img alt="" src="https://www.dtp-transit.jp/images/ss-688-986-72-20250713-213605.png" width="70%" />

### Features

- Create horizontal or vertical guides
- Specify position and margin (bleed)
- Unit selection (px, pt, mm)
- Auto get initial unit from preferences
- Increment/decrement values with up/down keys

### Flow

1. Show dialog
2. Input settings
3. Create guide on OK

### Change Log

- v1.0 (20250713): Initial version
- v1.1 (20250714): Added layer selection and repeat functionality
- v1.2.1 (20260706): Extracted dialog-window creation into a helper, added a locale-aware colon, appended colons to input labels, two-column layout with a bottom button bar, consolidated units into a shared dropdown + per-field unit labels, reverted the canvas ruler-origin offset (fixed misplacement from coordinate-space mismatch), deduplicated UI with helpers (addPanel/addColumnGroup/addLabeledField), added tooltips to key fields, trimmed unused symbol cases
- v1.2.2 (20260817): Restore the "_guide" layer on cancel (remove the empty layer this run created, put back an existing layer's lock state), warn only once when the active layer is locked, cap the repeat count at 1000, and handle H/V in the capture phase so the letter never lands in a numeric field. Also trimmed the header to an overview + README pointer, added README links and the article URL to the basic-info block, split user settings from layout constants, extracted colors / stroke widths / the layer name into constants, added JSDoc to every function, aligned names with the naming rules (`L()` → `getLabel()` and friends), split the preview and dialog-building code into smaller functions, and merged the duplicated target-switch and H/V handlers

