# new

[![Direct](https://img.shields.io/badge/Direct%20Link-new.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/new.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/new.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Creates a new layer, a new artboard or a new document from the selected objects.
Which one is created is chosen with radio buttons in the dialog.

### Main Features

- A dialog to choose what to create (layer / artboard / document)
- A Duplicate checkbox decides whether the original objects stay in place
- Layer: creates a new layer holding the selected objects
  - Avoids duplicate layer names ("New Layer", "New Layer 2", and so on)
  - Moves or copies while preserving the original stacking order as far as possible
  - Selects the objects on the new layer afterwards
- Artboard: creates a new artboard from the selected objects
  - Reads the existing row/column arrangement and reuses its pitch and column count when inserting
  - Shifts the following artboards and their artwork so the existing layout is preserved
  - Creates it at the same size as the active artboard and keeps the relative position of the objects
- Document: creates a new document from the selected objects
  - Copies the original file to make the duplicate, leaving the original document untouched
  - In the duplicate, everything except the selected objects and the current artboard is removed
  - Swatches, symbols and document settings carry over as they are
- Japanese and English localization

### Notes

- Duplicate applies only to Layer and Artboard. Document always keeps the original, so the checkbox is dimmed there.
- If no document is open, or nothing is selected, the script warns and stops.
- Layer creates the new layer at the very top of the Layers panel.
- The insertion point and direction for Artboard can be changed via `ARTBOARD_INSERT_AFTER_CURRENT` and `ARTBOARD_DIRECTION_AXIS`. Spacing is estimated from the existing arrangement, falling back to the preference value when there are fewer than two artboards.
- Artboard computes relative positions against the artboard the selection's center sits on; if it sits on none, the active artboard is used.
- Artboard never changes the object hierarchy (owning layer or group) when moving or duplicating; it only moves coordinates.
- Document runs only on a saved document. When the file is unsaved or has unsaved changes, the on-disk state and the on-screen state differ, so the script warns and stops.
- The Document duplicate is created next to the original file as `temp-<original filename>`.

### Article

https://note.com/dtp_tranist/n/xxxxxxxx

### Script info

- Version: v1.0.0
- First release: 2026-07-29
- Last updated: 2026-07-29
