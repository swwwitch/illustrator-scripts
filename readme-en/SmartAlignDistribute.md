# Lay out and distribute with an auto-detected direction

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAlignDistribute.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartAlignDistribute.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignDistribute.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Lays the selected objects out vertically or horizontally and distributes them at a given spacing.
The direction can be detected automatically, and alignment (left/right or top/bottom), the preview bounds (visible or geometric) and random reordering are all supported.
When a vertical layout includes text, the script duplicates and outlines it once for measurement while the dialog is open, caching the resulting height (and width where needed) so it does not re-duplicate on every preview.

### Original idea

John Wundes - Distribute Stacked Objects v1.1

https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design

https://gorolib.blog.jp/archives/77282974.html

### Script info

- Version: v1.3.7
- Last updated: 2026-10-04

### Update history

- v1.3.0 (20260927) : Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.3.1 (20260928) : The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.2 (20260928) : The button row is now built with the shared part. Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held). Clip groups are now measured by their mask (text masks and clip groups nested in a group included)
- v1.3.3 (20260929) : Dialog opacity changed to 98%
- v1.3.4 (20260930) : Fixed an error when running with characters selected by the Type tool
- v1.3.5 (2026-10-01) The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.3.6 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.3.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
