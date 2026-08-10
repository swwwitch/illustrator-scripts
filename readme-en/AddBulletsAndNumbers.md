# AddBulletsAndNumbers

[![Direct](https://img.shields.io/badge/Direct%20Link-AddBulletsAndNumbers.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AddBulletsAndNumbers.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBulletsAndNumbers.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Adds a bullet or a sequential number to the start of every line in the selected text frames.
Radio buttons switch between Bullet, Numbered list and None, and the settings are tuned while watching a preview. Reordering lines is supported as well.

When the dialog opens it infers the current type, symbol, number style and separator from the existing line markers, so you can carry on from what is already there.

- Bullet: choose a symbol (• - ● ○ ◎ ■ □ ◆ ◇)
- Numbered list: digits, white circled digits, black circled digits, ABC or abc, with a start number, zero padding and a separator (`.`, `:`, `|`)
  (white and black circled digits are laid out with a single tab stop, like bullets)
- Position: set the marker position, the body position and the alignment (left / center / right)
  (default tab stops are derived per type and symbol from the font size — digits 1.5/2.0, ABC/abc 1.0/2.0, bullets 1.2, •/- body 0.8, and so on)
- Marker formatting: font, style, scale and baseline shift, plus separate colours for the symbol/number and the separator
- Paragraph settings: leading, space after, and indent (including aligning the wrap position)
- Reordering: select several lines and move them up, down, to the top or to the bottom, or sort by name, by number, or by number descending
- Reset per frame: with several text frames selected, numbering restarts in each frame
- Reset: clears the scale, baseline, colour, indent and tab stops in one go

With several text frames selected, numbering runs top to bottom (left to right at the same height).

### Script info

- Version: v1.1.1
