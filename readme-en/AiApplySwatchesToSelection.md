# AiApplySwatchesToSelection

[![Direct](https://img.shields.io/badge/Direct%20Link-AiApplySwatchesToSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/AiApplySwatchesToSelection.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiApplySwatchesToSelection.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that applies swatches, or predefined colors, to the selected objects and text.

### Features

- Application unit selectable as object, character, word, line or paragraph
- Application order selectable as as-is, reversed, random or fully random
- "Random" shuffles the color order and repeats it; "fully random" draws for each target, so nothing repeats
- Live preview on every radio change
- The swatches selected when the palette opens are captured and used by name from then on

### Usage

1. Select the objects or text.
2. Run the script to open the palette.
3. Choose the unit and the order; the result updates live.

### Notes

- There is no Apply button: the state when the palette closes is what sticks. Use Cmd+Z to undo.
- Even a single selected swatch takes priority over the predefined colors.
- AiApplySwatchesToSelection-dialog.jsx is the modal-dialog version.

### Update History

- v1.8.0
