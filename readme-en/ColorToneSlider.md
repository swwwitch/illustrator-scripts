# Adjust brightness with a slider

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorToneSlider.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ColorToneSlider.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorToneSlider.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Adjusts the brightness of the selected objects with a slider. Positive values lighten, negative values darken.

<img alt="The Adjust Brightness dialog" src="../png/ss-826-314-144-20260917-043212.png" width="50%" />

### Features

- Slider range -50 to +50; the Amount field accepts -100 to +100 for amounts beyond the slider range
- Hold shift while dragging the slider for 10% steps
- Arrow keys step the Amount field by 1, or by 10 with shift
- R resets the adjustment
- Commits as a single undo step so the history stays clean
- Supports CMYK, RGB, grayscale, and spot colors (tint)
- Japanese / English UI

### Usage

1. Select the objects.
2. Run the script.
3. Adjust the brightness by dragging the slider or typing a value in the Amount field.
4. Check the preview and click OK. Cancel or Esc restores the original colors.

### Notes

- Gradients, patterns, and images (linked or embedded) are not adjusted.
- Groups and compound paths are traversed, and the objects inside them are adjusted.
- Text is handled per text frame. Text with mixed character formatting may be left unchanged.
- The preview is rebuilt through Illustrator's undo. Performing other operations while the dialog is open may leave the history in an unexpected state.

### Article

https://note.com/dtp_tranist/n/n88e33648b19a

### Update History

- v1.0.1 (2026-09-17) Added 10% snapping while dragging the slider with shift, removed "(R)" from the Reset label, revised the UI wording (title, slider end labels, Amount field, alerts), added tooltips, and reorganized internal naming and functions
- v1.0 (2025-12-28) Initial release
