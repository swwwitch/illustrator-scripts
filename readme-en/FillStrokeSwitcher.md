# FillStrokeSwitcher

[![Direct](https://img.shields.io/badge/Direct%20Link-FillStrokeSwitcher.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/FillStrokeSwitcher.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Adjusts fill and stroke on the selected objects in one pass.
- Seven processing modes are offered as radio buttons across two panels, Convert and Erase.
- The mode can be switched with live preview, and Cancel restores the original appearance.

### Main Features

- Convert modes
  - **Fill ↔ Stroke**: swaps fill and stroke within each selected object
  - **Fill → Stroke**: applies the fill color to the stroke (adds a stroke width of 1 when missing)
  - **Stroke → Fill**: applies the stroke color to the fill
  - **Swap Between 2 Objects**: exchanges the fill and stroke appearance between the two selected objects (enabled and preselected only when exactly two objects are selected)
- Erase modes
  - **Erase Fill** / **Erase Stroke** / **Erase Fill and Stroke**
- **Preview**: temporarily shows the result and restores the original appearance on Cancel
- Supports gradients, compound paths, and groups (groups and compound paths are processed recursively)
- Text (point text, area text, text on a path) is processed character by character, preserving partial text styling
- Restores the original selection after OK
- Reports the number of path / text / selection-restore failures with up to eight detail lines

### Workflow

1. Validate the document and the selection (one or two objects)
2. Snapshot the current appearance and show the mode dialog
3. With preview on, restore then reapply each time the mode changes
4. Commit with OK (applying the mode if preview was off), or restore the original appearance with Cancel

### Not Supported

- No open document (alert)
- An empty selection, or three or more objects selected (alert)
- Item types other than paths, compound paths, groups, and text (placed images, symbols, and so on) are left untouched
- "Swap Between 2 Objects" only handles PathItem, CompoundPathItem, and TextFrame

### note

- Based on a script by しぶやみゃむさん, with added features and refactoring.
- https://note.com/shibumi/n/n5229b4357dd3

### Update History

- v1.1.0: Current version
