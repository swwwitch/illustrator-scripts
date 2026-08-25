# ExpandGradient-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-ExpandGradient--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ExpandGradient-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandGradient-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Runs Object > Expand on the selection to break a gradient into a given number of steps. A simplified variant of ExpandGradient.jsx.

A temporary .aia action is generated, loaded and played internally, then unloaded and deleted.

### Usage

1. Select the object with the gradient.
2. Run the script.
3. Set the step count and click OK.

### Options

**Steps**

Defaults to 5; any integer of 2 or more. Up/Down adjusts by 1, Shift+Up/Down snaps to multiples of 10.

**Expand afterwards**

When on, Pathfinder Crop (Live Pathfinder Crop) and Expand Appearance (expandStyle) run in turn as post-processing.

### Update History

- v1.0.1 (2026-05-25)
