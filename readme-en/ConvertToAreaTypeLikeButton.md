# Create and adjust area type

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
- Left / right indent (linkable with the link icon) and outer spacing
- Justification and vertical alignment are always centered
- Preview is always on and reflects changes instantly
- Number fields step with the stepper buttons on their left or the arrow keys (to the next whole number, 1.5 → 2; Shift to the next multiple of 10; Option by 0.1)

### Notes

- Vertical centering uses a frame-alignment action that is loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.

### Update history

- v1.1.2 (20260928): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure. The button row is now built with the shared part
- v1.1.1 (20260928): Replaced the Link checkbox with a link icon. The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)

### Script info

- Version: v1.1.2
