# RectangleToArc

[![Direct](https://img.shields.io/badge/Direct%20Link-RectangleToArc.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/RectangleToArc.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArc.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Updated: 2026-05-19
- Converts the selected rectangle into an arc through three points: the bottom-left corner, the middle of the top edge, and the bottom-right corner
- The radius, center, start angle and end angle are derived from the rectangle's width and height, and the arc is built from Bezier segments of at most 90 degrees
- The original rectangle is deleted

### Main Features

- Only closed four-point paths whose edges are horizontal and vertical are processed (rotated rectangles, trapezoids, rhombuses and irregular shapes are skipped)
- The generated arc has no fill and a stroke
- When the original rectangle has a stroke, its colour and weight are carried over
- The conversion still runs when the original has no stroke
- Multiple selections are supported

### Script info

- Version: v1.0.1
