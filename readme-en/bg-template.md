# bg-template

[![Direct](https://img.shields.io/badge/Direct%20Link-bg--template.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/bg-template.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/bg-template.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Creates a rectangle the size of the current artboard, places it on a "bg-template" layer, marks that layer as a template and sends it to the back.

### Features

- Fills with K20 in CMYK documents
- Fills with #999999 in RGB documents
- Creates the "bg-template" layer, marks it as a template and sends it to the back

### Usage

1. Make the target artboard active.
2. Run the script.

### Notes

- Marking the layer as a template uses a dynamic action. The action definition has to be built from an array joined with `join("\n")`; the current implementation uses `'''`, which is a syntax error in ExtendScript.

### Update History

- v1.0 (2025-07-29)
