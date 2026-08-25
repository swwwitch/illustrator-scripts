# LogoGridMaker-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-LogoGridMaker--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/LogoGridMaker-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LogoGridMaker-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Generates construction lines and clear space for a logo, based on the visibleBounds of the selection.

The script is aimed at extracting a grid structure from letterforms.

### Features

- Horizontal lines: horizontal stretch, extra lines above and below, extension to the left. Line detection as none / automatic / horizontal segments / even division
- Vertical lines: vertical stretch, extra lines left and right, vertical division, extension upward. Vertical and diagonal elements can be extracted
- Shared settings and clear space: a unit is defined from the division count and used to build the clear space; layer name, stroke weight, guides and grouping are configurable
- Presets: 1x1 / auto / element / left / up-3 / clear space

### Usage

1. Select the logo, or whatever the construction should be based on.
2. Run the script.
3. Pick a preset, or set the options individually, and run it.

### Notes

- Turning on clear space disables the horizontal and vertical line panels entirely; stroke weight and guides are then disabled and grouping is forced on.
- The current UI state can be exported as JSON in an array-paste form.

### Update History

- v1.3.1 (2026-04-09)
