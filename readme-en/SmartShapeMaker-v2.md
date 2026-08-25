# SmartShapeMaker-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartShapeMaker--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/SmartShapeMaker-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartShapeMaker-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Creates regular shapes — circles, polygons, stars and Reuleaux-style forms — from a single dialog with a live preview.

### Main Features

- Specify number of sides (0 = Circle, 3/4/5/6/8, or custom with slider)
- Circle panel:
  - Superellipse option (only when sides = 0)
  - Superellipse shape control (exponent)
  - Anchor Points panel (2 / 3 / 4 / 5 / 6)
  - When Superellipse is ON: Rotate is forced OFF, Live Shape is forced OFF, and the Anchor Points panel is dimmed
- Star panel:
  - Star option + Pentagram option (side-by-side)
  - Inner radius input + 0-100 slider
  - Inner radius controls are dimmed when Star is OFF
  - When Pentagram is ON, Rotate is forced OFF
- Triangle direction options (Left / Right / Down) when sides = 3
- Width (size) panel with unit display
- Rotation panel:
  - Auto angle is used when Rotate is OFF (Circle = 45°, Polygon = 360/(sides*2))
  - When sides = 3 and Rotate is enabled, Triangle direction defaults to Down (60°)
  - Arrow-key editing supported
- Reuleaux-style option (odd-sided polygons only), with an adjustable appearance amount (0-200%) that resets to 100% when enabled
- Options panel: Live Shape conversion on finalize, and Split at Anchor Points (creates open stroked segments)
- Dialog opacity and position are restored within the current Illustrator session
- Preview does not pollute the Undo history; the final result can be undone in a single step
- View Zoom slider above OK / Cancel

### Keyboard Shortcuts

- `E` : Circle (0)
- `A` : Toggle Rotate

### Script info

- Version: v1.9
