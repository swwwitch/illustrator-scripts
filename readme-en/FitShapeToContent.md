# Create a backing shape behind text

[![Direct](https://img.shields.io/badge/Direct%20Link-FitShapeToContent.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/FitShapeToContent.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitShapeToContent.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Quickly creates and adjusts a backing shape that fits a text frame or a group. Padding and rounded corners are set in a dialog, and every change is previewed on the artboard right away.

### Features

- One text frame or group selected: a rectangle is created behind it automatically
- A text frame or group plus a shape selected: the shape is reused as the backing shape and centered on the content
- A group of one text frame and one shape selected: the group is kept intact and its shape is reused
- Padding (width and height) and rounded corners (radius, pill shape) are adjusted with a live preview

### Usage

1. Select the text or group (add a shape to reuse an existing one, or select a text+shape group).
2. Run the script.
3. Adjust the padding and corners, then click OK.

### Options

| Item | Behavior |
| --- | --- |
| Adjust Shape | When off, the shape keeps its size and is only centered on the content |
| Padding (Width / Height) | Space added around the content |
| Link | Keeps width and height at the same value |
| Radius | Corner radius; turn it off for square corners |
| Pill Shape | Sets the radius to half the height and caps both ends; the width is computed automatically |

Padding and radius are shown and entered in the document's ruler unit. The arrow keys step the values.

| Key | Step |
| --- | --- |
| Up / Down | ±1 |
| Shift + Up / Down | ±10 (snaps to multiples of 10) |
| Option + Up / Down | ±0.1 |

"Adjust Shape" starts on when the rectangle is created automatically (only a text frame or group selected), and off when an existing shape is reused.

### Notes

- Padding is measured from the path itself, so changing the stroke weight does not change the visible margin.
- When an existing shape is reused with "Adjust Shape" on, its appearance is cleared so that resizing does not distort it; fill, stroke, and opacity are carried over.
- Pill shapes are flattened with Live Pathfinder Add before being committed.
- Clipping groups are not supported for measurement. Release the clipping mask before running the script.
- The preview is undone and rebuilt on every change. Cancel or Esc restores the original state.

### Article

https://note.com/dtp_tranist/n/n6e4a6a2b175f

### Update History

- v2.0.3 (2026-08-31) Padding and radius now honor the ruler unit. A selected text+shape group keeps its group and original shape. Padding is measured from the path instead of the stroke
- v2.0.2 (2026-05-25)
