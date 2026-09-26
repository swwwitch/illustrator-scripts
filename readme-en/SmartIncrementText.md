# Duplicate text while incrementing its digits or letters

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartIncrementText.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/increment/SmartIncrementText.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartIncrementText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Finds the digits, letters, dates or times in the selected text frame and duplicates it downwards, incrementing the value each time. Everything you change in the dialog is shown as a live preview.

### Features

- Increments numbers (`01`, `2025`, …) and single letters (`A`–`Z` / `a`–`z`)
- For dates (`2026年9月20日` / `2026/09/20` / `2026.9.20`) you pick year, month or day; the calendar rolls over correctly and a weekday in brackets follows along
- For times (`19:00`) you pick hours or minutes; the clock rolls over at 24 hours
- When there is more than one candidate, radio buttons let you choose which one to increment (`Num1` / `Num2` / `Alpha1`, …)
- Start value override and zero padding (the width grows automatically when the last value needs another digit)
- Optionally merges the copies into a single text object on OK (leading = font size + gap)
- Arrow keys change the values (Shift for 10, Option for 0.1)
- Remembers the last dialog position

### Usage

1. Select the source text frame.
2. Run the script.
3. Set the number of copies, the step and the gap (the preview follows along).
4. Press OK.

### Settings

| Item | Description |
| --- | --- |
| Copies | How many copies to create, not counting the original |
| Step | Amount added at each step (negative values count down) |
| Gap | Space between copies, added on top of the font size |
| Target | Which part of the text to increment |
| Start value | Starts from the value you enter instead of the one in the original text |
| Zero pad | Pads with leading zeros to match the original width, widening it when the last value needs another digit |
| Merge into one text | Joins the copies into a single text object when you press OK |

### Notes

- Alphabet increment supports single letters (`A`–`Z`) only. `A1` works, but the `AB` in `AB1` and the `Ver` in `Ver1` are not increment targets.
- Use SmartRenumber.jsx to renumber an existing sequence instead.

### Article

- [【Illustrator】連番を一括生成するスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n5f25ed17b123)

### Update History

- v2.0.0
- v2.0.2 (20260920) : Renamed the fields to Gap / Start value / Merge into one text, fixed zero padding for negative values, dropped the increment-target row when there is nothing to choose, and disabled OK while the start value is invalid
