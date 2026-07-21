# AlignToArtboards

[![Direct](https://img.shields.io/badge/Direct%20Link-AlignToArtboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AlignToArtboards.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that aligns the selected objects to a chosen position on each artboard in a multi-artboard document.
- Pick one of nine anchor points in a 3x3 grid and set horizontal / vertical margins.
- Live preview while the dialog is open; Cancel restores the original positions.

### Main Features

- Two alignment bases
  - **All Artboards**: groups the selected objects by the artboard containing their center point, then aligns each group to the chosen position on that artboard
  - **Based on Active Artboard**: uses the selection on the active artboard as the reference and moves selections on the other artboards to the same relative position (objects on the active artboard are not moved)
- Nine target positions: Top-Left / Top-Center / Top-Right / Middle-Left / Center / Middle-Right / Bottom-Left / Bottom-Center / Bottom-Right
- Margin offsets objects inward from the corresponding edge; horizontal and vertical can be set separately, and "Linked" (on by default) mirrors the horizontal value to the vertical one
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
- Margins in Center mode, and both margins and the target panel in Based on Active Artboard mode (all disabled)

### Update History

- v1.1.2 (20260515): Current version
