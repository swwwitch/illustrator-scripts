# GradientFromFill

[![Direct](https://img.shields.io/badge/Direct%20Link-GradientFromFill.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/GradientFromFill.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GradientFromFill.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Creates a linear gradient on the selected filled objects, starting from their original fill color.

### Features

- Solid objects use their own fill color as the start
- For a single object already filled with a gradient, the start color is chosen from its stops, excluding black, white and transparent
- End color selectable as black, white, transparent, complementary or tint
- Angle selectable as 0, 30, 45, 60 or 90 degrees
- Separate gradients, reversing and preview are supported

### Usage

1. Select the objects.
2. Run the script.
3. Set the start color, end color and angle, then click OK.

### Notes

- With several objects selected, the whole start-color panel is disabled.
- Cancelling restores the original fill. The original fill and selection are also restored if the preview throws.
- Compound paths, objects nested in groups, and objects inside clipping groups are all handled.

### Update History

- v1.1.0
