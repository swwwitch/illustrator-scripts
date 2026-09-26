# Create an artboard-sized background, template it and send it back

[![Direct](https://img.shields.io/badge/Direct%20Link-bg--template.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/bg-template.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/bg-template.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Creates a rectangle the size of the current artboard, places it on a "bg-template" layer, marks that layer as a template and sends it to the back.

### Features

- Fills with K45 in CMYK documents (editable in the dialog)
- Fills with #999999 in RGB documents
- Creates the "bg-template" layer, marks it as a template and sends it to the back

### Usage

1. Make the target artboard active.
2. Run the script.

### Notes

- Marking the layer as a template uses a dynamic action.

### Update History

- v1.0 (2025-07-29)
- v1.0.2 (2026-09-27): Changing RGB with the arrow keys now updates the hex field. Fixed other keys re-rounding the value. Field-label colons now follow the UI language
- v1.1.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
