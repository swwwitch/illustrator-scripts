# IncrementDatesAndNumbers

[![Direct](https://img.shields.io/badge/Direct%20Link-IncrementDatesAndNumbers.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/IncrementDatesAndNumbers.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that increments or decrements dates, weekdays, times, sequence numbers, and numeric values inside the selected text.
- When a date is shifted, the weekday in parentheses follows automatically.
- The dialog shows the original and the result side by side while you choose the step value.

### Main Features

- **Increment / Value**: the amount to add or subtract (integers only, default 1). Arrow keys adjust the value (Shift = 10, Option = 0.1); 0 leaves the text unchanged
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
