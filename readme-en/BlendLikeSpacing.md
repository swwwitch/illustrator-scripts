# Gradually widen the spacing of a row of objects

[![Direct](https://img.shields.io/badge/Direct%20Link-BlendLikeSpacing.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/BlendLikeSpacing.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BlendLikeSpacing.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Rearranges a row of lines or objects so the gaps widen by a constant ratio, keeping both ends in place. Adjust the ratio with a slider and check the result in the live preview.

### Main Features

- The first and last objects stay put; the ones between them move only along the row.
- The ratio is the size of each gap relative to the previous one. At 130%, each gap is 1.3 times the one before it. 100% means even spacing.
- Change the ratio with the slider, the field, the stepper buttons or the arrow keys (100–300%).
- Choose the direction on the Vertical or Horizontal row.
  - Down, Up, Right, Left: the gaps widen toward that side.
  - Center: the gaps are narrowest at both ends and widen toward the middle.
  - Ends: the gaps are narrowest in the middle and widen toward both ends.
  - Random: the gaps vary at random, and the ratio is the widest gap over the narrowest. Click Shuffle for a new pattern.
- Raise Count under Objects above the number selected to add copies of the top object (the leftmost in a horizontal row), up to 200. With two objects selected, the count starts at three.
- The dialog box opens with Down for a vertical row and Right for a horizontal row.
- With objects of different sizes, the space between objects follows the ratio.
- A single undo reverts the result.

### Usage

1. Select two or more lines or objects in a row.
2. Run the script.
3. Choose a direction and adjust the ratio with the slider.
4. Click OK to confirm, or Cancel to return everything to its original position.

### Notes

- A group moves as one object.
- Bounds follow the Use Preview Bounds preference.
- When objects overlap so much that there is no space between the two ends, the distance between their leading edges is divided by the ratio instead.
- Directions along an axis where the objects are perfectly aligned cannot be selected.

### Update History

- v1.2.0 (2026-11-09): Added Random to the directions and Count under Objects
- v1.1.0 (2026-11-09): Added Center and Ends to the directions
- v1.0.0 (2026-11-06): Initial version
