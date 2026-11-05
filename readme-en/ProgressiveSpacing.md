# Gradually widen the spacing of a row of objects

[![Direct](https://img.shields.io/badge/Direct%20Link-ProgressiveSpacing.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/ProgressiveSpacing.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ProgressiveSpacing.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Rearranges a row of lines or objects so the gaps widen by a constant ratio, keeping both ends in place. Adjust the ratio with a slider and check the result in the live preview.

### Main Features

- The first and last objects stay put; the ones between them move only along the row.
- The ratio is the size of each gap relative to the previous one. At 130%, each gap is 1.3 times the one before it. 100% means even spacing.
- Change the ratio with the slider, the field, the stepper buttons or the arrow keys (100–300%).
- Choose the direction the gaps widen toward: Down, Up, Right or Left. The dialog box opens with Down for a vertical row and Right for a horizontal row.
- With objects of different sizes, the space between objects follows the ratio.
- A single undo reverts the result.

### Usage

1. Select three or more lines or objects in a row.
2. Run the script.
3. Choose a direction and adjust the ratio with the slider.
4. Click OK to confirm, or Cancel to return everything to its original position.

### Notes

- A group moves as one object.
- Bounds follow the Use Preview Bounds preference.
- When objects overlap so much that there is no space between the two ends, the distance between their leading edges is divided by the ratio instead.
- Directions along an axis where the objects are perfectly aligned cannot be selected.

### Update History

- v1.0.0 (2026-11-06): Initial version
