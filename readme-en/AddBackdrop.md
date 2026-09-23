# Lay a shape behind text

[![Direct](https://img.shields.io/badge/Direct%20Link-AddBackdrop.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/AddBackdrop.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBackdrop.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Generates a shape sized to the visual bounds of the selected text (or objects) and places it behind them.

An existing backdrop is detected and replaced.

<img alt="The Add Backdrop dialog" src="../png/ss-990-1102-144-20260923-190826.png" width="50%" />

### Features

- Targets point text and area text; a non-text selection uses the whole selection bounds
- Shape selectable as circle, superellipse or rectangle (keys E / S / R)
- Margins set separately or linked for the vertical and horizontal axes (default: a quarter of the short side)
- Undo-based preview that commits in one step on OK

### Usage

1. Select the text or objects that need a backdrop.
2. Run the script.
3. Choose the shape and margins, then click OK.

### Article

https://note.com/dtp_tranist/n/na8af4a7016ad

### Update History

- v1.6.4 (2026-09-23): Revised UI wording (title, panel names, label colons, tooltips); split the button row and added a Show Transparency Grid button
- v1.6.3 (2026-09-23): Fixed typed stroke weight and CMYK values not being clamped, the previous corner radius not being restored when rounding was on, and the message shown with no document open. Internal cleanup
- v1.6.1 (2026-03-26)
