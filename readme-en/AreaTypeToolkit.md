# AreaTypeToolkit.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-AreaTypeToolkit.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AreaTypeToolkit.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AreaTypeToolkit.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

Headings and button labels often start life as point text, and only later do you wish they were Area Type — so the copy wraps inside a box when it grows, or sits centred in it.

Doing that by hand means drawing a rectangle, converting it, pouring the text in, and setting the font again. And once you have Area Type, tuning it (frame size, leading, indents, text placement) still means moving between panels.

This script puts **creation and adjustment into one flow**.

<img alt="" src="" width="50%" />

## Usage

1. Select point text, path text, a shape (closed path), or existing Area Type
2. Run the script
3. Pick a creation method in the convert dialog, then set the typesetting in the adjust dialog

**With only Area Type selected, the convert dialog is skipped and you start at the adjust dialog.**

With nothing selected the script warns and stops.

## Convert dialog

Four creation methods. Methods that don't apply to the current selection are dimmed automatically.

| Method | What it does |
| --- | --- |
| Simple | Uses a rectangle the size of the point text plus 1pt |
| Button style | Uses a rectangle 1.2x wide and 1.8x tall, with the text centred both ways |
| Use selected object | Duplicates the selected shape as the frame and pours the selected text into it |
| Dummy text on selected object | Turns the selected shape itself into Area Type and fills it with dummy text |

| Selection | Available methods |
| --- | --- |
| Point / path text only | Simple, Button style |
| Shapes only | Dummy text on selected object |
| Text plus shapes | Use selected object, Dummy text on selected object |

Path text is split off into point text first, automatically. Per-character font, size, fill, stroke, and leading are snapshotted and restored, so the appearance survives the trip.

### Pairing text with shapes

"Use selected object" accepts several text-and-shape pairs at once. Each text is matched to **the shape whose bounds contain the text's centre**, or failing that **the shape whose centre is nearest**. Matching uses `geometricBounds` (rectangular bounds).

If nothing could be converted, the script says so and leaves the dialog open.

## Adjust dialog

Everything about the Area Type's typesetting, in one place. Preview is on by default.

### Separate text

Breaks Area Type apart into **a rectangle plus point text**.

| Option | What it does |
| --- | --- |
| No adjust | Keeps the Area Type and applies the settings below (default) |
| Stroke 1pt black | Separates, and gives the rectangle a 1pt black stroke |
| No path | Separates, and leaves the rectangle unpainted |
| Remove path | Separates, and deletes the rectangle |

Choosing anything but "No adjust" dims the other panels.

### Font size

| Control | What it does |
| --- | --- |
| Font size | Set directly |
| Make overset | Shrinks the font until the overset clears (binary search, up to 40 passes) |
| Fit | Grows until it oversets, then shrinks — filling the frame |

"Make overset" and "Fit" are unavailable when the frame has more than one paragraph.

### Leading

Leading is set as an **auto-leading amount (%)**. Illustrator shows it as "Auto", so the ratio holds when the font size changes.

| Control | What it does |
| --- | --- |
| Leading | Auto-leading amount, in % |
| Effective | Font size x %, in points. Enter a value here to back-calculate the % |

For text with a fixed leading the fields open empty, and nothing is applied while they stay empty.

### Justification / Text alignment

Justification (left, center, right, justify with last line left, justify all lines) and text alignment within the frame (top, center, bottom, justify).

Justification also responds to the **L / C / R / J / F** keys (disabled while an input field has focus).

Text alignment cannot be set reliably through the DOM, so it is applied via a dynamic action (`adobe_frameAlignment`) loaded from a temp file at startup and discarded on exit — nothing is left in the Actions panel.

### Frame size

Width and height in the ruler unit. The width can also be driven by a character count (Japanese UI only — Roman glyph widths vary too much for the arithmetic to hold).

[Auto] toggles Area Type auto-sizing, also via a dynamic action (`adobe_SLOAreaTextDialog`). It **does not run during preview**, since calling `app.doScript` from a modal dialog is unstable.

### Indent / Options

- Left and right indents ("Link" keeps them equal)
- Spacing between the frame and the text
- Leader tabs — sets a right-aligned tab with a "…" leader at 400pt on every paragraph. Turning it on switches justification to right

## Input fields

Every numeric field responds to the arrow keys.

| Key | Step |
| --- | --- |
| Up / Down | ±1 |
| Shift + Up / Down | ±10 (snaps to multiples of 10) |
| Option (Alt) + Up / Down | ±0.1 |

## Targets

TextFrame (point / path / area text), PathItem and CompoundPathItem (closed paths)

## Notes

- Width and height reject zero, negative, NaN, and extreme values, falling back to the last valid entry (the cap is the equivalent of 100000pt).
- Preview is undone with `app.undo()`. Mixing it with other operations may leave things out of step.
- Ruler units supported: mm, cm, inch, pt, pica, px, Q.

## Change log

- v1.2.0 (2026-07-28) Added the Leading panel. Fixed reading justification and indents from existing Area Type. "Use selected object" now handles multiple pairs. Added Q (ha) ruler unit. Added a warning when nothing can be converted
- v1.1.3 (2026-03-04) Added input validation for width and height
- v1.0.0 (2026-03-03) Initial release
