# Distribute colors across objects and text

[![Direct](https://img.shields.io/badge/Direct%20Link-AiApplySwatchesToSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/AiApplySwatchesToSelection.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiApplySwatchesToSelection.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A modal dialog that applies swatches, or predefined colors, to the selected objects and text.

### Features

- Application unit selectable as object, character, word, line or paragraph
- Application order selectable as as-is, reversed, random or fully random
- "Random" shuffles the color order and repeats it; "fully random" draws for each target, so nothing repeats
- Live preview on every radio change
- Colors come either from the swatches selected when the dialog opens, or from a swatch group picked in the dropdown
- Units that do not fit the selection are dimmed automatically
- "Per word" staggers colors so each line starts on a different color

### Usage

1. Select the objects or text.
2. Select the swatches to use in the Swatches panel.
3. Run the script, then choose the unit and the order; the result updates live.
4. [OK] commits the result; [Cancel] or Esc reverts it.

### Notes

- Even a single selected swatch takes priority over the predefined colors.
- When no swatches are selected, auto colors are used (CMYK: two-channel CM/CY/MY mixes; RGB: six predefined colors).
- The swatch group dropdown skips the unnamed (uncategorized) group and any group with no colors; it is dimmed when the document has no groups.
- The preview is reverted from a snapshot of the original fill, stroke and opacity rather than with `app.undo()`, which would roll back the whole document history.
- Per-character previewing colors the first 500 characters only; the rest is colored when you press [OK].

### Update History

- v1.8.1 (2026-09-09): Merged the persistent-palette and modal-dialog versions into the dialog version
- v1.8.0 (2026-07-19): Preview is now reverted from a snapshot baseline; per-character coloring of long text is decimated while previewing
- v1.7.3 (2026-07-17): Lighter behavior while the UI is open
- v1.7.2 (2026-07-17): Added "fully random" to the application order
