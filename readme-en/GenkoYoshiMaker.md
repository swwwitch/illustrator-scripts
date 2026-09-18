# Set text on manuscript paper

[![Direct](https://img.shields.io/badge/Direct%20Link-GenkoYoshiMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/GenkoYoshiMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GenkoYoshiMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

This script matches the character advance of the selected text to a grid of cells, draws a rule at every character interval, and adds rules along both sides of each line. Vertical and horizontal text are both supported, and the writing direction is left as it is.

A dialog sets the rule density, the extra cells and lines, the rule styles and the crosshairs. The result is previewed on every change, so you can check it without closing the dialog.

### Main features

- Character attributes matched to the grid (scale, tracking, tsume, proportional metrics, aki, kerning)
- The scale is set in the dialog (90% by default); turning it off keeps the scaling but still matches the advance to it
- A quarter of the tracking added above the first character as kerning, centering the glyphs in their cells
- Both vertical and horizontal text (the panel names follow the writing direction)
- Horizontal rules at one-character intervals, from the top of the first character to the bottom of the last
- Square cells the size of the font, so the scaled-down glyphs do not narrow the grid
- Vertical rules along both sides of the text, with an adjustable extension
- Side rules a quarter of the font size outside the grid
- Thicker horizontal rules at a fixed character interval
- Solid or dashed crosshairs at the center of every cell; the dash and the gap are equal and divide one cell
- Empty cells added before and after the text
- Empty lines added left and right of the text, spaced by a quarter of a cell
- Rule density set with a K0-K100 slider (Shift snaps to steps of 10%)
- Presets that fill the whole dialog at once, with a Save button that adds the current settings
- The settings confirmed with OK come back as the defaults until Illustrator quits
- Live preview that stays open while you adjust the settings
- Rules created on a dedicated layer (`Manuscript Grid`) and sent to the back
- Japanese and English UI

### How to use

1. Select a single vertical text object.
2. Run `GenkoYoshiMaker.jsx`.
3. Pick a preset, or set the options in the Overall, Characters, Vertical rules, Horizontal rules and Crosshairs panels.
4. Check the result in the preview (turn off "Preview" to hide it).
5. Click "OK" to draw the rules.

### Options

| Option | Default | Description |
| --- | --- | --- |
| Preset | Standard | Standard, No side rules, No crosshairs, Simple or Custom; editing any value switches to Custom |
| Rule density | K50 | Density of the rules, set with a slider (Shift snaps to steps of 10%) |
| Extra cells | 0 | Empty cells added before and after the text |
| Extra lines | 0 | Empty lines added left and right of the text, spaced by a quarter of a cell |
| Adjust the character scale | On | Scale the characters to the cells; the advance is matched either way |
| Scale | 90% | Horizontal and vertical scale of the characters |
| Extension | A quarter of the font size | How far the line rules run past the cells; the default is a quarter of the font size in millimeters, rounded to one decimal |
| Add an outer rule | On | Add vertical rules a quarter of the font size outside the outermost line |
| Horizontal rules | Solid | Draw the horizontal rules solid or dashed (1mm / 1mm) |
| Thicker every n characters | On / 5 | Thicken the horizontal rule at the given interval (0.25mm) |
| Crosshair style | Dashed | None, solid or dashed |
| Save | - | Adds the current settings to the preset list and shows the code for them |
| Crosshair segments | 9 | Number of dashes per crosshair line; the dash and the gap are equal and divide one cell |
| Preview | On | Show the result without closing the dialog |

Numeric fields step with the up and down arrow keys (Shift steps by 10 and snaps to multiples of 10).

### Rule specifications

| Rule | Weight | Density |
| --- | ---: | ---: |
| Rules along each line | 0.25mm | Rule density |
| Cell rules | 0.1mm | Rule density |
| Emphasized cell rules | 0.25mm | Rule density |
| Outer rules | 0.25mm | Rule density |
| Crosshairs | 0.1mm | 30% (`LAYOUT.crossGray`) |

In vertical text the line rules run vertically and the cell rules horizontally; in horizontal text it is the other way around, and the dialog panels swap accordingly.

In RGB documents, the equivalent gray RGB values are used.

### Character attributes

The following attributes are set so that the characters land on the grid.

| Attribute | Value |
| --- | --- |
| Tracking | Derived from the scale (`100000 / scale - 1000`, about 111 at 90%) |
| Tsume | 0 |
| Proportional metrics | Off |
| Aki before and after | Auto |
| Horizontal and vertical scale | The scale set in the dialog (90% by default) |
| Kerning (first character only) | A quarter of the tracking (1/1000 em) |
| Justification | Start of the writing direction (top for vertical, left for horizontal); the text is moved back so it does not shift |
| Leading | One cell plus a quarter of a cell |
| Automatic kerning | Japanese monospaced |

### Notes

- Compatible with Illustrator 2024 to 2026.
- Select exactly one text object before running the script.
- The cell size comes from the font size of the first character, so text with mixed font sizes will not line up with the rules.
- Tracking is measured in 1/1000 em, so it does not change with the font size; it is recalculated only when the scale changes.
- The kerning on the first character is applied after the grid is measured, so that the rules do not shift down with the text.
- With several lines, each line is spaced by a quarter of a cell and the leading is set to match, so the original leading is not kept.
- Running the script again rebuilds only the `Rules` group drawn for that text; grids drawn for other text objects are left alone.
- The preview temporarily shows a copy of the selected text. Cancelling restores the original state.
- Automatic kerning is set to Japanese monospaced, and manual kerning is applied to the first character only.
- The line rules and the outer rules are drawn with projecting caps.
- Half-width characters do not fill one cell, so the text drifts from the rules when they are mixed in.
- After the run, the script reports only the attributes that could not be set.

### Change log

- v1.0.0 (2026-09-19): Initial release
