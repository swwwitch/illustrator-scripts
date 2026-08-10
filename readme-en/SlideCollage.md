# SlideCollage

[![Direct](https://img.shields.io/badge/Direct%20Link-SlideCollage.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SlideCollage.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SlideCollage.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Places selected .ai / .pdf files (PDFs by page) on the active document in a grid to build a portfolio-style thumbnail sheet.

- Import: place by artboard number (for example `1-20` or `1,3,5`)
- Item: choose the PDF crop box (art / trim / bleed / bounding) and apply a corner radius (converted to points)
  - Each item is clipped into a group by a same-sized rectangle, and the corner radius (a live effect) is applied to that clip group
- Grid: set direction (horizontal / vertical / random), column count and spacing
- Even columns: a distribution mode adds one extra slot to even columns, and their vertical offset can be tuned separately
- Layout: scale (an extra multiplier on top of the auto-fit result), rotation (rotates the whole sheet and centers it on the artboard) and position offset (horizontal / vertical)
- Mask: clips inside the margins on OK, with an adjustable mask corner radius applied to the clip group
- Artboard: an optional background colour (entered as HEX) can be added, and it is excluded from the mask

Most changes are reflected in a real-time preview (120 ms debounce); changing the artboard numbers on import is the exception.
Numeric fields step with Up/Down, Shift and Option.

### Original idea

Slide Collage - Portfolio Layout Generator -

https://slide-collage.vercel.app/

### Script info

- Version: v1.5
