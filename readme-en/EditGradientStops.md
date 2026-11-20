# List and edit gradient stops

[![Direct](https://img.shields.io/badge/Direct%20Link-EditGradientStops.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/EditGradientStops.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditGradientStops.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Lists the gradient stops used by the selected objects and lets you change their color, location and opacity, or swap stops.

### Features

- Collects fill and stroke gradients, including those inside groups, compound paths and text, and lists them in one panel per gradient
- Each stop is one row: color chip, color values (CMYK, RGB, gray, Lab, or swatch name and tint), location and opacity, in location order
- Click a color chip to change the color with Illustrator's standard color picker
- Change location and opacity with the ∧∨ buttons, the arrow keys or by typing (arithmetic allowed), from 0 to 100%
- The ↑↓ buttons at the end of a row swap color and opacity with the stop above or below (locations stay)
- The duplicate button at the end of a row adds a copy of the stop right below it (placed halfway to the next stop, or to the previous stop on the last row)
- The delete button at the end of a row removes the stop (at least two stops remain)
- Preview (Option+A, on by default) shows the changes on the objects right away
- Gradients with identical stops share one panel, and changes go to every fill and stroke that uses them

### Usage

1. Select objects that use gradients.
2. Run the script.
3. Change the color, location and opacity, then click OK. Cancel restores the original stops.

### Notes

- Gradients in fills or strokes added in the Appearance panel are not covered.
- For text, the fill and stroke returned for the whole text frame's character attributes are read and written.
- The color picker returns colors in the document's color mode, so picking a new color for a swatch color turns it into a process color.
- Up to 10 gradients are listed (change MAX_LISTED_GRADIENTS at the top of the script).
- Whether Preview starts on is set by DEFAULT_PREVIEW_ON at the top of the script.

### Article

[List and edit gradient stops (Japanese)](https://note.com/dtp_tranist/n/n524fafa198dd)

### Update History

- v1.0.1 (2026-11-20) Clicking the script name and version at the bottom left opens About This Script (copy the script name, open the README and other links, show in Finder; shared part updated)
- v1.0.0 (2026-11-20) First release
