# Sort numbers, letters or Japanese numerals and renumber them in sequence

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartRenumber.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SmartRenumber.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartRenumber.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Sorts the selected text — digits, letters, or Japanese numerals — in a chosen order and renumbers it as a sequence. Every change in the dialog shows up in the preview right away.

### Main features

- Radios under Base Value pick the format (`123` / `ABC` / `abc` / `一二三` / `I II III` / `壱弐参`), dropping `1`, `A`, `a`, `一`, `I`, or `壱` into the start value
- Typing straight into the start value switches the format too (`C` gives C, D, E…)
- Letters carry over from `Z` to `AA`, `AB`, and so on
- Japanese numerals are composed as `十`, `二十一`, `百三`, `一万`, and so on
- Roman numerals run `I, II, III, IV`…, written in capitals however they are typed
- Formal Japanese numerals run `壱, 弐, 参, 四 … 九, 拾, 拾壱`…, using the four legally distinct characters 壱弐参拾
- The start value can be `0` or negative (`-3`)
- Six sort orders: current value, horizontal, vertical, Z-pattern, N-pattern, and stacking order
- Reverse turns the order around before renumbering
- The width typed into the start value is always kept (`01` → 01, 02, 03)
- Zero Padding widens the sequence to the widest number, and is dimmed when the width would not change
- Adjust Stacking (on OK) rearranges the stacking to follow the new numbers; it sits under the sort orders and starts checked
- Prefix and Suffix add text before and after the number
- The start value defaults to the smallest value in the selection, in its own format
- Up/Down keys step the value (digits: Shift for ±10 snapped to a multiple of ten, Option for ±0.1; letters and Japanese numerals step one at a time)
- Tooltips on every option
- Dialog with a live preview (Cancel restores the original text)
- Japanese / English UI

### Usage

1. Select the text objects.
2. Run `SmartRenumber.jsx`.
3. Set the start value under Base Value and pick the sort order, checking the preview as you go.
4. Click OK to apply.

### Settings

| Option | Description |
| --- | --- |
| Format (123 / ABC / abc / 一二三 / I II III / 壱弐参) | The format to write; choosing one drops its lowest value into the start value |
| Start Value | The first value of the new sequence: a number (`1` / `0` / `-3`), a letter (`A` / `a`), a Japanese numeral (`一` / `壱`), or a Roman numeral (`I`) |
| Sort Order | The order the numbers are assigned in (see below) |
| Reverse | Turns the order around before renumbering |
| Zero Padding | Widens the sequence to the widest number (digits only); the typed width is kept either way |
| Adjust Stacking (on OK) | Reorders the stacking to follow the new numbers, lowest number frontmost. Checked by default |
| Prefix | Text placed before the number |
| Suffix | Text placed after the number |

### Sort orders

| Sort order | Description |
| --- | --- |
| Current Value Order | Ascending order of the current values; letters go alphabetically and Japanese numerals by their value |
| Horizontal (Left to Right) | From the leftmost object to the right |
| Vertical (Top to Bottom) | From the topmost object down |
| Z-Pattern (Left-to-Right, Row-major) | Left to right within a row, then down to the next row |
| N-Pattern (Top-to-Bottom, Column-major) | Top to bottom within a column, then right to the next column |
| Stacking Order (Front to Back) | From the frontmost object backward |

### Notes

- Only text objects made up entirely of digits, entirely of letters, or entirely of Japanese numerals are picked up. Mixed contents such as `No.1`, `Item 3`, or `A1` are ignored, so a run that added a prefix or suffix cannot be re-run on its own result; strip it back to a bare number first.
- Full-width digits (`１`) are not supported.
- Letters such as `I` and `X` read as either letters or Roman numerals, so the `I II III` radio decides which one applies.
- The formal numerals share `四`-`九`, `百`, `千` and `万` with the plain ones, so the `一二三` / `壱弐参` radio decides which style continues.
- Existing Roman numerals sort as letters under Current Value Order; sorting by position is unaffected.
- Adjust Stacking is not reflected in the preview: the stacking looks the same on screen, so it runs on OK only.
- Cancelling or closing the dialog writes the recorded original strings back.
- With Adjust Stacking and a selection spanning several layers or groups, each container is reordered on its own.
- For the Z-pattern and N-pattern, objects within 10 pt of each other count as the same row or column.
- Positions are compared using each text object's top-left corner (`left` / `top`).
- Use SmartIncrementText.jsx to duplicate while generating a new sequence.

### Article

- [Renumber selected text with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/nf3b6601cd165)

### Changelog

- v2.1.0 (2026-09-21): Added letters (A, B, C / a, b, c), Japanese numerals (一, 二, 三) Roman numerals (I, II, III) and formal numerals (壱, 弐, 参), and renamed Start Number to Start Value. Added an Options panel holding Reverse and Zero Padding, plus a new Adjust Stacking option under the sort orders (checked by default). Moved the start value into a Base Value panel with format radios (123 / ABC / abc / 一二三 / I II III / 壱弐参). Rebuilt the dialog as two columns with the buttons underneath. The width typed into the start value is now always kept, and Zero Padding is dimmed when the width would not change. Fixed the value being rewritten by keys other than Up/Down, arrow keys filling an empty field, Stacking Order following the selection order instead of the stacking order, and ragged zero padding for a negative start value. Replaced the app.undo() preview rollback with restoring the recorded strings, fixing text that could vanish while the start value was being edited. Reorganised the internal naming and structure
- v2.0 (2026-01-09)
