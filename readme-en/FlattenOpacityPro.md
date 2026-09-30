# Bake opacity into the fill color

[![Direct](https://img.shields.io/badge/Direct%20Link-FlattenOpacityPro.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/FlattenOpacityPro.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenOpacityPro.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Bakes the opacity of the selected objects into their fill colors so that everything becomes fully opaque.
When the selection contains gradients, images, or other items that cannot be baked, it falls back to the built-in Flatten Transparency.

### Features

- Parent group opacity is composited recursively
- Overlapping objects are composited from the back so the apparent color is reproduced
- When the selection contains anything that cannot be baked (gradient, pattern, gray, or spot colors; text, images, symbols, and the like; or a blending mode other than Normal), the built-in Flatten Transparency is applied to the whole selection
- The blend method can be switched between linear-light RGB and the source color space (`USE_GAMMA_CORRECT_BLEND`)

### Usage

1. Select the objects.
2. Run the script.

### Notes

- The tolerance for treating shapes as identical is set by `GEOM_TOL_PT` (position and size) and `AREA_TOL` (area).
- Effects (such as drop shadows) cannot be detected, so such objects are still baked.
- The change cannot be undone cleanly, so duplicating the file first is recommended.

### Update History

- v1.0
- v1.0.2 (2026-09-27) Alerts are now shown in English as well
- v1.1.0 (2026-09-29) Merged FlattenTransparency.jsx; selections containing anything that cannot be baked are now processed with the built-in Flatten Transparency
