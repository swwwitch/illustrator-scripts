# FlattenOpacityPro

[![Direct](https://img.shields.io/badge/Direct%20Link-FlattenOpacityPro.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/FlattenOpacityPro.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenOpacityPro.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Bakes the opacity of the selected objects into their fill colors so that everything becomes fully opaque.

### Features

- Parent group opacity is composited recursively
- Overlapping objects are composited from the back so the apparent color is reproduced
- The blend method can be switched between linear-light RGB and the source color space (`USE_GAMMA_CORRECT_BLEND`)

### Usage

1. Select the objects.
2. Run the script.

### Notes

- The tolerance for treating shapes as identical is set by `GEOM_TOL_PT` (position and size) and `AREA_TOL` (area).
- The change cannot be undone cleanly, so duplicating the file first is recommended.

### Update History

- v1.0
