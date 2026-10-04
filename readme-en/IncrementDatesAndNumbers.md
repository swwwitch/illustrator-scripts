# Duplicate text while incrementing dates and numbers

[![Direct](https://img.shields.io/badge/Direct%20Link-IncrementDatesAndNumbers.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/increment/IncrementDatesAndNumbers.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/IncrementDatesAndNumbers.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that increments or decrements dates, weekdays, times, sequence numbers, and numeric values inside the selected text.
- When a date is shifted, the weekday in parentheses follows automatically.
- The dialog shows the original and the result side by side while you choose the step value.

### Main Features

- **Step / Amount**: the amount to add or subtract (integers only, negatives allowed, default 1). The ∧∨ buttons to the left of the field and the Up/Down arrow keys change it by 1 (Shift-click / Shift+arrow snaps to the next multiple of 10); 0 leaves the text unchanged
- **Target**: shown only for texts containing a year/month/day, using the actual values as radio labels (e.g. "2025", "11", "21"). Choose which part to shift (day by default); switching resets the step value to 1
- **Type**: shown only for two-part dot patterns like "12.1", to choose between Number (default) and Date interpretation
- The original and the result are previewed live at the top of the dialog (numeric cases also show the computed value)
- Supported string patterns
  - Japanese dates: 2025年11月21日, 2025年11月21日（金）, 11月21日㊎
  - Era-based dates: 令和7年3月21日（金）, etc. (the era name is updated when the shift crosses an era boundary)
  - Slash / dot dates: 2025/11/21, 2025/11/21(Fri), 2025.11.21, 29.3.2, 29.3
  - 8-digit dates: 20251121
  - Times: 19:00 (the "year" target shifts hours, others shift minutes, with minute carry / borrow and 24-hour wrapping)
  - Standalone weekdays: 日–土, Sun–Sat, circled symbols ㊐–㊏
  - Standalone years: 2025
  - Generic numbers: 123, 1,234, 100.5 (decimals keep their precision and shift the last digit; thousands separators are preserved)
- Processes multiple selected text frames and text inside groups; multi-line text is handled line by line
- Automatic Japanese / English UI

### Workflow

1. Analyze the first text frame in the selection to detect the pattern (year/month/day, two-part dot, etc.) and build the dialog accordingly
2. Set the step value, target, and type — the result preview updates on every change
3. On OK, walk the selection and shift the date, weekday, time, or number line by line
4. Write the updated strings back into the text frames

### Not Supported

- Decimal step values (integers only)
- Objects other than text frames (groups are traversed recursively for text)
- Only the first matching pattern per line is processed (additional dates or numbers on the same line are left untouched)

### Update History

- v1.2 (20251118): Public release
- v1.2.2 (20260930): Added stepper buttons (∧∨) to the amount field; the arrow keys now step the same way as the buttons (whole numbers only). Layout, button row, and dialog position/opacity now use the shared parts. Renamed the dialog to "Increment Dates and Numbers" and the panel/field to "Step" / "Amount:", and revised the tooltips for Type, Target, and the result
- v1.2.3 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.2.4 (2026-10-01): Removed the hand-written dialog centering; position and opacity now come from the shared part. Unified the window and panel margins and spacing with the shared layout part
- v1.2.5 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.2.6 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
