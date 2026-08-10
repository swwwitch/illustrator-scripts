# AdjustFontSize

[![Direct](https://img.shields.io/badge/Direct%20Link-AdjustFontSize.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AdjustFontSize.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that adjusts the font size and horizontal / vertical scale of the selected characters.
- The result is shown as a live preview, and Cancel restores the state from before the dialog opened.
- "Reset" discards the adjustments, and option-clicking it unifies a string with mixed sizes to the size of its first character.
- A range selected with the Type tool targets exactly those characters; a text object selected with the Selection tool targets all of its characters.

### Main Features

- Font Size Adjustment
  - Font Size: entered in the ruler unit (the text unit from Preferences)
  - Scale: one value sets both horizontal and vertical scale (%)
  - Apparent: the actual visual size computed as font size × scale (dimmed while the scale is 100%)
  - Arrow keys increment / decrement the value (shift for steps of 10, option for 0.1 — 5 for the scale)
- "Actual ↔ Apparent" button converts between size and scale
  - Each press toggles between baking size × scale into the actual font size at 100% and restoring the previous scaled state
  - Typing a size or scale by hand discards the saved pre-bake state
- "Reset" discards the adjustments and returns to the just-opened state
  - Option-click unifies every selected character to the first character's font size and sets the horizontal / vertical scale to 100%
- On launch the fields show the actual size and scale of the first selected character
- Supports point text and area text, multiple selections, and range selections made with the Type tool
- Automatic Japanese / English UI

### Workflow

1. Read the selected characters and open the dialog with the first character's size and scale as the initial values
2. Apply the font size and scale to the selected characters, previewing on every change
3. OK commits (as a single undo entry); Cancel restores the state from before the dialog opened

### Not Supported

- No open document (the script exits silently)
- No text selected (an alert is shown and the script exits)
- Baseline shift, kerning, tracking, character alignment and auto kerning (use AdjustTextScaleBaseline)

### Update History

- v1.0.0 (20260802): Initial release
