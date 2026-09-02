# Scatter confetti across the selected area

[![Direct](https://img.shields.io/badge/Direct%20Link-ConfettiMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/ConfettiMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConfettiMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Scatters confetti across the selected object, or across the active artboard when nothing is selected. Shape, count, distribution and randomness are tuned against a live preview in the dialog, and OK commits the result onto a `Confetti` layer.

### Features

- Pick any combination of shapes: circle, rectangle, square, triangle, star, sparkle A, sparkle B, heart, ribbon and symbol
- Three distributions — uniform, top to bottom, and radial outward — each with an adjustable strength
- Randomize size, opacity, skew and rotation, with a slider for each amount
- Mask option that clips the confetti to the shape of the selected object
- Margin option that expands the generation area outward
- Live preview in the dialog, plus a slider for the canvas zoom
- OK commits the result onto a `Confetti` layer
- Japanese / English UI

### Usage

1. Select the object that defines the area to scatter over (leave everything deselected to use the artboard).
2. Run the script.
3. Choose the shape and amount, check the preview, and click OK.

### Options

**Basic**

| Item | Description |
| --- | --- |
| Base Size | Reference size of a single piece of confetti |
| Count | Number of pieces to generate (10–500) |
| Mask | Clip the confetti to the shape of the selected object |
| Margin | How far the generation area extends outward; the upper bound is derived from the target size |
| Distribution | Uniform, top to bottom, or radial outward. Strength controls how strong the bias is |

**Shapes**

Each piece picks randomly from the shapes you check.

- **Option (Alt) + click**: turn on only that shape
- **Cmd + Option (Alt) + click**: turn off only that shape

Soloing circle, sparkle B or heart with Option + click also turns rotation (and skew) off, so the outline stays readable.

**Randomize**

| Item | Description |
| --- | --- |
| Size | Variation around the base size (100 keeps it fixed, 300 allows up to ±150%) |
| Opacity | The further right the slider goes, the more faded pieces appear |
| Skew | Maximum shear angle about the center (0–45°) |
| Rotation | Maximum random rotation (0–360°) |

### Notes

- Turning Margin on switches Mask off automatically; the two are mutually exclusive.
- Mask is unavailable when a text frame is selected, and when the artboard is used as the target.
- Symbol can only be chosen when the document already contains symbols.
- The output is committed exactly as the preview shows it — nothing is regenerated on OK.
- With Mask off the result is collected into a single group; with Mask on it is emitted as a clipping group.

### Article (Japanese)

https://note.com/dtp_tranist/n/n5a41fb524a5a

### Update History

- v1.7.4 (2026-09-03) Internal cleanup (naming, structure, table-driven shape selection) that made the preview lighter. Fixed the preview not refreshing when Margin was switched off
- v1.7.3 (2026-03-11)
- Initial release (2026-02-16)
