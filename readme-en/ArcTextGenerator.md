# ArcTextGenerator

[![Direct](https://img.shields.io/badge/Direct%20Link-ArcTextGenerator.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ArcTextGenerator.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that generates an arc-shaped path sized to the selected point text / path text and converts it into type on a path.
- The curve depth, direction, and Type on a Path effect are set in a dialog, with live preview.
- The generated arc path has no fill and no stroke (invisible), so only the text is visible.

### Main Features

- **Curve:** slider (0-100, default 50) for the arc depth. 0 is almost straight; 100 gives the roundest curve
- **Direction:** the direction the arc bulges — Up (default) or Down
- **Fit to path width:** None (default) / Font size / Tracking
  - Font size: change the font size so the text reaches both path endpoints (shrinks in 0.1 pt steps, growing first when needed)
  - Tracking: keep the font size and adjust letter spacing with a coarse pass followed by a fine pass
- **Effect:** Type on a Path effect — Rainbow (default) / Skew / 3D Ribbon / Stair Step / Gravity
- **Tracking:** when the checkbox is on, adds the given value to the existing tracking (-100 to 500; arrow keys adjust, Shift = 10, Option = 0.1). Turning it off resets it to 0. Dimmed while "Fit to path width: Tracking" is selected
- **Preview:** on by default. Shows a temporary result while enabled; turning it off or cancelling restores the original
- Paragraph justification is always centered
- Automatic Japanese / English UI, with tooltips on every option

### Workflow

1. Collect point text / path text from the selection (recursing into groups)
2. Measure the real text size via temporary outlines and create a straight path along the baseline
3. Bend the straight path into an arc by moving its Bézier handles, create the type on a path, and duplicate the original text ranges
4. Apply center justification, tracking, and the effect, then fit to the path width with the chosen method (any path selected together with the text is removed)

### Not Supported

- No open document
- Selections containing no point text or path text (area type is not a target)
- Empty text, or text whose lines / text ranges cannot be read
- Fitting is skipped for text on closed paths and for locked, hidden, or non-editable text

### note

- Original idea: Toshiyuki Takahashi (@gautt) https://note.com/gautt/n/n92f6faeda048

### Update History

- v1.1.0 (20260519): Public release
