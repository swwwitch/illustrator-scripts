# 4lines

[![Direct](https://img.shields.io/badge/Direct%20Link-4lines.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/outline/4lines.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/4lines.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Analyzes the paths of outlined text and draws the four typographic lines — descender, baseline, mean line and ascender.

### Usage

1. Select the outlined text.
2. Run the script.

### Notes

- The estimate comes from a histogram of anchor points and horizontal segments. If there are not enough peaks, the script warns and exits.
- Compound paths — glyphs with counters such as `o` or `A` — are handled by their compound bounding box.
- Text that has not been outlined has no PathItem, so it cannot be processed.

### Update History

- v1.0
