# Step the font size of text, or the stroke weight of other objects

[![Direct](https://img.shields.io/badge/Direct%20Link-AdjustFontSizeOrStroke.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AdjustFontSizeOrStroke.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSizeOrStroke.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Steps the font size of selected text, or the stroke weight of other selected objects, by a fixed amount in the units set in Preferences. It runs immediately without a dialog box, so it is meant to be assigned to a keyboard shortcut.

### Features

- Text: increases the font size by 1 in the Type unit under Preferences > Units (1 pt for points, 1 Q for Q)
  - Characters with different sizes each get the step added to their own size
  - With characters selected by the Type tool, only those characters change
- Other objects: increases the stroke weight by 0.1 in the Stroke unit under Preferences > Units (0.1 pt for points, 0.1 mm for millimeters)
  - Paths without a stroke are left alone
- Goes inside groups and compound paths, handling text and paths in the same selection
- Rounds the result in the unit, so repeated runs leave no stray decimals

### Options

Set these in the "User Settings" block at the top of the script.

| Variable | Values | Meaning |
|---|---|---|
| `STEP_SIGN` | `1` / `-1` | `1` to increase, `-1` to decrease |
| `STEP_MULTIPLIER` | `1` / `5` / `10` | Step multiplier (x1, x5, x10) |
| `FONT_SIZE_STEP` | `1` | One font size step (in the Type unit) |
| `STROKE_STEP` | `0.1` | One stroke weight step (in the Stroke unit) |

To use separate shortcuts for increasing, decreasing or each multiplier, duplicate the file, change the settings, and assign each copy its own shortcut.

### Notes

- The font size stays within 0.1–1296 pt and the stroke weight within 0.1–1000 pt (never 0 pt, so a thinned stroke does not become invisible).
- Character strokes on text are not changed (text only gets its font size changed).
- Strokes added in the Appearance panel are not changed (only the basic stroke is).
- Locked or hidden objects inside groups are left alone.
- With fixed (non-auto) leading, larger text may look cramped.

### Update History

- v1.0.0 (2026-11-01): Initial release
- v1.0.1 (2026-11-01): Raised the minimum stroke weight from 0 pt to 0.1 pt, so a thinned stroke does not become invisible
