# Fit the view so objects fill a given share of the window (reusable template)

[![Direct](https://img.shields.io/badge/Direct%20Link-FitViewToItems.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/_templates/FitViewToItems.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitViewToItems.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A reusable template that centers the view on objects and sets the zoom so they fill a given percentage of the window.

It can add a "Fit to Window [65] %" row to a dialog, and includes helpers to restore the original view on Cancel.

### Key features

- Centers the view on the objects' bounds (effects excluded) and zooms in or out so they fill the given share
- Keeps the zoom within the range Illustrator accepts (3.125% to 6400%)
- Adds a "Fit to Window [%]" row with built-in Japanese / English labels and tooltips
- `captureView` / `restoreView` save and restore the view position and zoom

### How to use

1. Copy the whole `var FitViewToItems = (function () { ... })();` block into the target script's IIFE.
2. Save the view before the dialog opens, and add the row.

        var initialView = FitViewToItems.captureView(doc);
        var fitViewControls = FitViewToItems.addControls(optionPanel, { value: false, percent: 65 });

3. Call it after the preview has been rebuilt.

        if (fitViewControls.checkbox.value) {
            FitViewToItems.fit([previewItem], { doc: doc, fillRatio: fitViewControls.getFillRatio() });
            app.redraw();
        }

4. Keep the field's enabled state in step with the checkbox, and restore the view on Cancel.

        fitViewControls.checkbox.onClick = function () { fitViewControls.updateEnabled(); refitView(); };
        /* on Cancel */
        FitViewToItems.restoreView(initialView, doc);

### Notes

- Zoom changes are not recorded in the undo history, so the caller has to run `restoreView` on Cancel.
- Refitting on every setting makes the zoom jump around; call it only when the size changes (such as the width) and when the dialog opens.
- To leave the view alone when the objects are already visible, use [KeepInView](KeepInView.md) instead.
- Scripts stay self-contained without `#include`, so copy the block rather than including it.
- Used in: `jsx/shape/SmartShapeMaker.jsx` ("Fit to Window")

### Changelog

- v1.0.0: Initial release (template, extracted from SmartShapeMaker v2.3.0)
