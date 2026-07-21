# SmartSelectionFilter

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartSelectionFilter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SmartSelectionFilter.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that narrows down the current selection, keeping only the text frames and paths that match the chosen conditions.
- Conditions cover text kind, paragraph alignment and font, plus path open/closed state and fill / stroke status.
- Every checkbox change is reflected on the canvas without closing the dialog.
- Cancel restores the selection and visual state from before the script ran.

### Main Features

- "Text" panel: Point Text / Area Text / Path Text (each shown with its count)
- "Alignment" panel: Left / Center / Right (disabled when no text kind is checked)
- "Font" panel: lists every font used in the selection so it can be filtered per font (with counts; empty when no text is present)
- "Path" panel: Open Path / Horizontal Line / Vertical Line (open paths whose height or width is nearly zero are treated as horizontal / vertical lines)
- "Closed Path" panel: Fill Only / Stroke Only / Fill and Stroke
- "Quick Select" panel: Text / Stroke Only / Fill Only in one click. Changing the detailed options keeps these checkboxes in sync automatically
- "Non-target Objects" panel: Do Nothing / Hide / Opacity (0–100% slider, current value shown in the radio label). The slider is enabled only while Opacity is chosen
- Option (Alt) clicking a checkbox sets every checkbox to the same state at once
- Button to toggle between Outline and Preview display mode (the original mode is restored on exit)
- Automatic Japanese / English UI

### Workflow

1. Check that a document is open and something is selected, then store the selection as an array
2. Walk the selection (recursing into groups) and count items per kind, alignment, font and path state
3. On each option change, recompute the selection and the non-target display state and apply it to the canvas immediately
4. Commit the current selection with OK, or restore the original selection and visual state with Cancel

### Not Supported

- No open document, or nothing selected
- Objects other than text frames and paths (group contents are processed recursively, but groups themselves never remain selected)
- Closed paths with neither fill nor stroke

### Update History

- v1.0.0: Initial release
