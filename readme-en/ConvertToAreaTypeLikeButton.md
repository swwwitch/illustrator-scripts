# ConvertToAreaTypeLikeButton

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToAreaTypeLikeButton.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ConvertToAreaTypeLikeButton.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToAreaTypeLikeButton.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A tool to create and adjust area type from point text, path text, shapes, or existing area type.

Depending on the selection it converts to area type automatically and opens the adjust dialog.

- Point / path text only → convert to button style (width ×1.2, height ×1.6)
- Text + shape → fill the shape (as area type) with the text
- Area type only → open the adjust dialog directly

### Adjust dialog

- Set font size, or auto-shrink to the largest fitting size via "Make overset"
- Change frame size (width / height)
- Left / right indent (linkable) and outer spacing
- Justification and vertical alignment are always centered
- Preview is always on and reflects changes instantly

### Notes

- Vertical centering uses a frame-alignment action that is loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.

### Script info

- Version: v1.0.0
