# SwapObjects

[![Direct](https://img.shields.io/badge/Direct%20Link-SwapObjects.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SwapObjects.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script to swap the positions of two selected objects in Illustrator.
- Even when the objects have different widths, the outer left and right edges of the pair can be kept in place.
- The reference can be switched with live preview, and Cancel restores the original positions.

### Main Features

- Two position references
  - **Swap Center Positions**: the objects exchange their center points (the outer edges shift when widths differ)
  - **Swap Keeping Outer Edges**: the outer left and right edges stay fixed, so the total span is unchanged
- Switchable size reference
  - Off: path geometry (geometric bounds)
  - On: appearance including strokes and effects (visible bounds)
- Live preview (Cancel restores the original positions)
- Defaults to "Swap Center Positions" for equal widths, "Swap Keeping Outer Edges" otherwise
- Automatic Japanese / English UI, with tooltips on every option

### Workflow

1. Validate that exactly two swappable objects are selected (locked or hidden items are excluded)
2. Choose the position and size references in the dialog
3. Snapshot the original positions and bounds, calculate the offsets, and apply them
4. Commit with OK, or restore the original positions with Cancel

### Not Supported

- Selections other than exactly two objects
- Objects that cannot be moved or whose bounds cannot be read
- Locked or hidden objects, including those inside a locked or hidden layer or group

### note

https://note.com/dtp_tranist/n/na534a676fae2

### Update History

- v1.2.1 (20260406): Public release
- v1.3.0 (20260721): Fixed "Swap Keeping Outer Edges" doing nothing, refactored the whole script (IIFE, shared layout helpers, categorized labels), added tooltips and a no-document check
