# Align the selection to one of nine points on each artboard

[![Direct](https://img.shields.io/badge/Direct%20Link-AlignToArtboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AlignToArtboards.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AlignToArtboards.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that aligns the selected objects to a chosen position on the artboard.
- Pick one of nine anchor points in a 3x3 grid and set horizontal / vertical margins.
- Live preview while the dialog is open; Cancel restores the original positions.

### Main Features

- Two alignment bases
  - **All Artboards**: groups the selected objects by the artboard containing their center point, then aligns them to the chosen position on that artboard **one object at a time**
  - **Based on Active Artboard** (default): aligns the selection on the active artboard to the chosen position **while keeping its internal layout**, then places selections on the other artboards at the same relative position
- In a single-artboard document there is nothing to distribute, so All Artboards is dimmed and the base stays on Based on Active Artboard
- Nine target positions: Top-Left / Top-Center / Top-Right / Middle-Left / Center / Middle-Right / Bottom-Left / Bottom-Center / Bottom-Right
- Margin offsets objects inward from the corresponding edge; horizontal and vertical can be set separately, and "Linked" (on by default) mirrors the horizontal value to the vertical one
- The target and the margin apply to both alignment bases; the margin is disabled only when the target is Center
- Margin input follows the Illustrator ruler unit (rulerType) and is converted to points internally
- "Use Preview Bounds" switches between visual bounds including strokes and effects (on) and geometric bounds of the shape (off)
- Keyboard shortcuts
  - Target: q=Top-Left / w=Top-Center / e=Top-Right / a=Middle-Left / s=Center / d=Middle-Right / z=Bottom-Left / x=Bottom-Center / c=Bottom-Right
  - Alignment base: 1=All Artboards / 2=Based on Active Artboard
  - Arrow keys step the margin fields (Shift=±10, Option=±0.1); Enter/Return triggers OK
- Automatic Japanese / English UI, with tooltips on every option

### Workflow

1. Check that a document is open and something is selected
2. Choose the alignment base, target position, margins, and bounds type in the dialog (the preview re-applies on every change)
3. Group the selection by the artboard containing each center point and compute the offset to the chosen anchor
4. Commit with OK; Cancel or closing the dialog reverts the preview translation

### Not Supported

- No open document, or an empty selection
- Locked or hidden objects, including those inside a locked or hidden layer or group
- Objects whose center point falls outside every artboard, or whose bounds cannot be read
- Based on Active Artboard with nothing selected on the active artboard (there is no reference, so nothing moves)
- Margins when the target is Center (disabled)

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n50aacdeb4908)

### Update History

- v1.1.2 (20260515): Current version
