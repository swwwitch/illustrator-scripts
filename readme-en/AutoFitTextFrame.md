# AutoFitTextFrame

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoFitTextFrame.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AutoFitTextFrame.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that resolves overset text in selected area type / path text by adjusting the font size or the frame height.
- Besides simply removing the overset, it can also maximize the font size up to the largest size that still fits.
- In documents using variables / data sets, the original values are recorded in a tag and reset before each data set is processed.

### Main Features

- **Shrink Text to Fit** (on by default): shrink the font size in 0.1 pt steps until the overset is gone
- **Maximize Text Size** (on by default): grow the font size (doubling) until it oversets, then shrink to fit — so the text fills the frame even when there was slack
- With both on, "Maximize" runs first, then "Shrink"
- **Adjust Area Text Height**: available only when the selection contains area type. Turning it on disables the font-size options and enables:
  - **Adjust height**: toggle Auto Size on then off, expanding the frame just enough and fixing that height
  - **Auto size**: apply Auto Size to the area type (expand only; never turned back off)
- When manual leading is set, the leading follows the font size at the same ratio
- Area type uses Illustrator's built-in `overflows` check when available; path text uses the character count on visible lines
- Text inside selected groups and text ranges (cursor selections) are also processed
- Remembers the dialog position; automatic Japanese / English UI

### Workflow

1. Alert and exit when nothing is selected; otherwise collect area type / path text recursively from the selection (de-duplicated)
2. On the first data set, record the original font size / height in a tag, and reset from it before each run
3. Apply the chosen processing (Maximize / Shrink / Adjust height / Auto size) to the targets
4. Remove the tags after the last data set is processed

### Not Supported

- No open document, or nothing selected
- Point text (only area type and path text are targets)
- Locked, hidden, or non-editable text
- Text containing line breaks (the font-size processing alerts and aborts)
- Height adjustment and Auto Size do nothing when the selection contains no area type

### Update History

- v2.3.1 (20260304): Public release
