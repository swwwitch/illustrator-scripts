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
  - Down, Right: the gaps widen toward that side.
  - Center: the gaps are narrowest at both ends and widen toward the middle.
  - Random: the gaps vary at random, and the ratio is the widest gap over the narrowest. Click Shuffle for a new pattern.
  - Turn on Reverse to swap the wide and narrow sides: Down becomes Up, Right becomes Left, and Center becomes Ends (narrowest in the middle, widening toward both ends). The radio button labels change to match. Reverse is not used with Random.
- Enter a number in Add under Number of Objects to add that many copies of the top object (the leftmost in a horizontal row). Use the field or the slider; it starts at 0 and goes up to 64. Total beside it shows the selected objects plus the copies.
- Turn on Fill or Stroke under Color Blend to step the fill or stroke colors of the objects between the two ends.
  - In groups, the objects inside are paired from the front.
  - Two tints of the same spot color blend by tint. Different color types blend in the document color mode.
  - When either end has no fill (or stroke), a gradient, or a pattern, that attribute is left unchanged.
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

### Article

[【Illustrator】並んだ線の間隔をだんだん広くするスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n67c10b92e5fd)

### Update History

- v1.6.1 (2026-11-09): Renamed the dialog box to Blend-Like Spacing
- v1.6.0 (2026-11-09): Folded Up, Left and Ends into a Reverse option, leaving three radio buttons per row
- v1.5.0 (2026-11-09): Count is now the number to add (Add, 0 to 64), with Total shown beside it
- v1.4.0 (2026-11-09): Added Color Blend (fill and stroke)
- v1.3.0 (2026-11-09): Added a slider to Count; renamed the panel to Number of Objects; lowered the Count limit to 100
- v1.2.0 (2026-11-09): Added Random to the directions and Count under Objects
- v1.1.0 (2026-11-09): Added Center and Ends to the directions
- v1.0.0 (2026-11-06): Initial version
