# RightMarkPlacer.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-RightMarkPlacer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/RightMarkPlacer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RightMarkPlacer.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script scans the selected objects from left to right and places a mark in the middle of the gap between each adjacent pair.

Use it to show the flow or the relationship between elements in a layout.

## Main features

- Nine marks to choose from (`▶` / `>` / `>>` / `─` / `→` / `➡` / `─\` / `＋` / `×`)
- Places the same mark between every adjacent pair in the selection
- The width is derived from the narrowest gap, on the safe side (it can also be typed in)
- A width you type is the drawn size of the mark
- Height (%), stroke width and position (horizontal / vertical) are adjustable
- Position accepts negative values and can be stepped with the arrow keys
- Follows the current ruler units (and the stroke unit setting for the stroke width)
- `▶` can be reshaped with Inset (up to 80% of the width) and rounded with Rounded corners
- `─\` takes an angle for the slash, supports round ends, and can be mirrored
- `>` / `>>` switch to a filled shape with Keep top and bottom edges horizontal (both draw the same width)
- `➡` is sized from the stroke width alone; its head is three times the stroke width tall
- The preview updates immediately and is always cleaned up when the dialog closes
- Japanese and English UI

## Usage

1. Select two or more objects.
2. Run `RightMarkPlacer.jsx`.
3. Choose the shape and set the options (turn on Preview to check the result).
4. Click OK to run.

## Shape panel

| Mark | Description |
| --- | --- |
| `▶` | Filled triangle. Inset dents its left edge; Rounded corners rounds it. |
| `>` | Stroked chevron. Keep top and bottom edges horizontal turns it into a filled shape. |
| `>>` | Two stroked chevrons. Gap sets the distance between them. Supports the horizontal option. |
| `─` | Horizontal line. |
| `→` | Stroked arrow (shaft plus head). |
| `➡` | Solid arrow. The shaft is drawn at the stroke width, then outlined and merged with the head. |
| `─\` | Horizontal line plus a slash. Angle sets the slope of the slash. |
| `＋` | Cross. |
| `×` | Diagonal cross (the cross rotated 45°). |

- Mirror horizontally flips the generated mark (disabled for `─` / `＋` / `×`).
- For `>>`, Width is the width of a single chevron; the pair spans width × 2 + gap.
- For `×` alone, Width is the arm length shared with `＋`, so the drawn width is about 0.707 × the value.

## End Style panel

Chooses None or Round for the stroke caps and joins. It is disabled for the fill-only shapes (`▶` / `➡`).

## Options panel

| Item | Description |
| --- | --- |
| Height | Percentage of the combined height of the two adjacent objects, up to 200%. Unused for `─` / `＋` / `×` / `➡`. |
| Width | Drawn width of the mark. Leave it empty to derive it from the narrowest gap. |
| Inset | The dent in the left edge of `▶`, up to 80% of the width. |
| Gap | Distance between the two chevrons of `>>`. Negative values are allowed. |
| Stroke | Stroke width. The minimum is the equivalent of 0.25 pt, and the alert states it in the unit on screen. Unused for `▶`, which is fill-only. |
| Angle | Slope of the slash in `─\`, up to 89°. |
| Keep top and bottom edges horizontal | Draws `>` / `>>` as a filled shape with horizontal top and bottom edges. |
| Rounded corners | Rounds the corners of `▶` (with or without an inset). |

- Items that do not apply to the current shape are dimmed, including their label and unit.
- Typing a width stops the automatic calculation; clearing the field turns it back on.
- Height, width, stroke, angle and position all return to the defaults of the shape you switch to.

### Sizing of `➡`

`➡` is sized from the stroke width alone. Height is unused.

| Stroke width | Shaft | Head height | Head depth |
| --- | --- | --- | --- |
| 3.6 pt (default) | 3.6 pt | 10.8 pt | 5.4 pt |
| 5 pt | 5 pt | 15 pt | 7.5 pt |

- The head is three times the stroke width tall (`ARROW3_HEIGHT_TO_STROKE_RATIO`).
- Its depth is half its height, capped at 90% of the width (`MAX_ARROW_HEAD_RATIO`).
- `→` caps the depth of its head the same way, so the head never exceeds Width, while its height still follows Height.

## Position panel

| Item | Description |
| --- | --- |
| Horizontal | Shifts the mark left or right. Negative values are allowed. |
| Vertical | Shifts the mark up or down. Negative values are allowed. |

- Both reset to 0 when the shape changes.

## Keyboard

| Key | Action |
| --- | --- |
| `F` | Set the end style to None |
| `R` | Set the end style to Round |
| `V` | Toggle Mirror horizontally |
| `↑` / `↓` | Step the focused field by one step |
| `Shift` + `↑` / `↓` | Step while snapping to multiples of ten steps |
| `Option` (`Alt`) + `↑` / `↓` | Step by a tenth of a step |

- `F` / `R` / `V` do nothing while a modifier (`command` / `control` / `option` / `shift`) is held, so shortcuts such as `command` + `V` still work.
- The step size follows the unit: `1` for percent, degrees, pt, px, mm and Q/H, `0.1` for cm, and `0.01` for inches.

## Settings

These can be changed in the User settings section at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `MARK_COLOR_CMYK` | `[0, 0, 0, 100]` | Mark color (CMYK) |
| `MAX_HEIGHT_PERCENT` | `200` | Maximum height percentage |
| `MIN_STROKE_WIDTH_PT` | `0.25` | Minimum stroke width in points |
| `MAX_INSET_RATIO` | `0.8` | Maximum inset as a ratio of the width |
| `TRI_CORNER_RADIUS_RATIO` | `0.12` | Rounded-corner radius ratio for `▶` |
| `SLASH_ANGLE_DEFAULT` / `SLASH_ANGLE_MAX` | `35` / `89` | Slash angle and its maximum |
| `ARROW3_HEIGHT_TO_STROKE_RATIO` | `3` | Height of the head of `➡` as a multiple of the stroke width |
| `MAX_ARROW_HEAD_RATIO` | `0.9` | Maximum head depth of `→` / `➡` as a ratio of the width |
| `AUTO_WIDTH_GAP_RATIO` | `0.7` | Ratio of the gap used for the automatic width |
| `AUTO_WIDTH_GAP_RATIO_SMALL` | `0.35` | Ratio used for the automatic width of `＋` / `×` |
| `AUTO_WIDTH_CHEVRON_HEIGHT_RATIO` | `0.5` | Ratio of the height used for the automatic width of `>` / `>>` |

## Notes

- The script alerts and exits when no document is open or fewer than two objects are selected.
- It also alerts and exits when the active layer is locked or hidden.
- **No mark is created for a pair whose objects overlap or leave no gap.** When no pair could be filled, the alert appears before the dialog opens.
- The script assumes a left-to-right arrangement; objects are sorted by the X coordinate of their center.
- Text is measured from its outlined bounds (the measurement copy is removed automatically). The measurements are kept until the dialog closes, so text is not re-outlined on every preview refresh.
- Marks are created on the current layer and left selected.
- Clicking OK while Preview is on keeps the previewed objects as the result.
- `➡` uses the Outline Stroke and Unite menu commands, so the selection changes during processing and is restored afterwards.
- The number of decimals shown follows the unit: two for pt and mm, more for inches and centimeters.

## Article

[DTP Transit Annex | note (Japanese)](https://note.com/dtp_tranist/n/nebac730ec187)

## Update history

- v1.3.2 (2026-08-01): Unified the shape-creation code; reorganized the label definitions and layout settings. Made Width the drawn size for every shape (fixing the mismatch between the stroked and filled `>` / `>>`), kept the heads of `→` / `➡` inside Width, sized `➡` from its stroke width (head height = stroke × 3, dropping the Stroke ×3 label), added unit-aware decimals and arrow-key steps, and moved the locked-layer and no-gap checks to startup
- v1.3.1 (2026-04-04): Fixed negative position values when stepped with the arrow keys
- v1.0.0 (2026-03-28): Initial version
