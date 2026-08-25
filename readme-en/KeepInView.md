# KeepInView

[![Direct](https://img.shields.io/badge/Direct%20Link-KeepInView.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/_templates/KeepInView.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KeepInView.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A reusable template that scrolls the view only when the objects you created or changed have fallen outside it.

Objects already in view are left alone, so the canvas does not jump on every operation.

### Features

- Scrolls only when the target has moved off screen
- Zooms out — never in — when the target does not fit, so the zoom level never creeps up on its own
- Can add a checkbox whose label and tooltip are already localized in Japanese and English

### Usage

1. Copy the whole `var KeepInView = (function () { ... })();` block into your script's IIFE.
2. Add the checkbox. The wording is built in, so you do not have to supply any.

        var cbKeepInView = KeepInView.addCheckbox(grpViewOptions, { value: true });

3. Call it once you have finished creating the result.

        if (cbKeepInView.value) KeepInView.ensureVisible(createdItems, { doc: doc, fitRatio: 0.9 });

### Notes

- If you build the checkbox yourself, skip `addCheckbox` and call `ensureVisible` only.
- Omitting `doc` uses the frontmost document; omitting `fitRatio` uses `DEFAULT_FIT_RATIO`.
- `#include` is deliberately avoided so that each script stays a single file — copy the block instead.
- Example: `jsx/text/DynamicTextGenerator.jsx` ("Keep the result in view")

### Update History

- v1.0.0: First release (template)
