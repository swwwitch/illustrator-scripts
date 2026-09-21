# A reusable color-picker library

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/ColorPicker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A reusable color-picker library meant to be included from other scripts.

### Usage

1. Include it with `#include "ColorPicker.jsx"` at the top of the calling script (outside any function).
2. Call `ColorPicker.show()`.

        var result = ColorPicker.show({
            value: "FF0000",      // "RRGGBB" or "cmyk:C,M,Y,K"
            title: "Color Picker",
            lang: "en"            // label language ("ja" or "en"; defaults to "en")
        });

3. It returns `null` when the dialog is cancelled.

### Notes

- Running it on its own opens nothing.
- If the `#include` sits inside a function, this file's `SCRIPT_NAME` / `SCRIPT_VERSION` and the like become variables of that function and hide the caller's values of the same name.
- Included by: `jsx/shape/SmartShapeMaker.jsx` / `jsx/stroke-table/LeaderLineBuilder.jsx`
- `jsx/text/AddBulletsAndNumbers.jsx` has had the picker built in since v1.2.2 and no longer includes this file.

### Update History

- v1.0.2 (2026-09-21): Fixed an error on opening the picker, caused by the sliders not receiving the tooltip language
- v1.0.1 (2026-09-19): Added tooltips to the controls
- v1.0
