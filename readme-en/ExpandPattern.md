# Expand pattern fills into regular paths

[![Direct](https://img.shields.io/badge/Direct%20Link-ExpandPattern.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ExpandPattern.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandPattern.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Groups the selected objects, applies 3D Rotate (Classic) and Pathfinder Merge effects, then expands the appearance. Pattern fills and the like become regular paths without changing how they look.

### Usage

1. Select the objects to expand.
2. Run the script.

No dialog is shown; the script runs right away.

### What it does

1. Groups the selection and names the group "Expanded Pattern"
2. Applies these effects to the group (listed below "Contents" in the Appearance panel, in this order)
   - 3D Rotate (Classic): Position "Front" (X/Y/Z all 0°), Surface "No Shading"
   - Pathfinder Merge
3. Runs Object > Expand Appearance

### Notes

- The effects are applied as LiveEffect XML. The parameters are based on the defaults in [live-effect-functions-for-illustrator](https://github.com/mark1bean/live-effect-functions-for-illustrator).
- Merge applied from the menu lands above "Contents", so the script stacks it below "Contents" through XML instead.

### Update History

- v1.0.0 (2026-10-04) Initial release
