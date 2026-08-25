# AlignLabelWidths

[![Direct](https://img.shields.io/badge/Direct%20Link-AlignLabelWidths.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/_templates/AlignLabelWidths.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AlignLabelWidths.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A reusable template that measures the rendered width of ScriptUI statictext labels and aligns them all to the widest one.

It keeps the colon and the value column lined up in stacked "label: value" information or settings panels.

### Features

- Measures the actual rendered width, so it does not depend on the locale or on how long the wording is
- A simpler variant based on a fixed `characters` count is included as well
- Ships with a helper that adds a label and a value as a single row

### Usage

1. Copy `alignLabelWidths()` (or the simpler `setLabelsFixedWidth()`), and `addLabelValueRow()` if you need it, to the top level of your script.
2. Collect the labels into an array as you build each row.
3. Pass the collected array to the alignment function.
   - Option A (recommended, measured automatically): `alignLabelWidths(labels, 'right')`
   - Option B (simple, fixed character count): `setLabelsFixedWidth(labels, 13, 'right')`

### Notes

- More robust than a fixed `characters` value, but it has to be called once the layout is settled.
- This is a template; running it on its own does nothing.

### Update History

- v1.0.0: First release (template)
