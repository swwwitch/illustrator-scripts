# ColorPicker

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/ColorPicker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A reusable color-picker library meant to be included from other scripts.

### Usage

1. Include it with `#include "ColorPicker.jsx"`.
2. Call `ColorPicker.show()`.

        var result = ColorPicker.show({
            value: "FF0000",      // "RRGGBB" or "cmyk:C,M,Y,K"
            title: "Color Picker"
        });

3. It returns `null` when the dialog is cancelled.

### Notes

- Running it on its own opens nothing.
- Example: `jsx/shape/SmartShapeMaker.jsx`
- `jsx/text/ColorPicker.jsx` is an identical file.

### Update History

- v1.0
