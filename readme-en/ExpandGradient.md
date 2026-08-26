# ExpandGradient

[![Direct](https://img.shields.io/badge/Direct%20Link-ExpandGradient.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ExpandGradient.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandGradient.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Runs Object > Expand on the selection to break a gradient into a given number of steps.

A temporary .aia action is generated, loaded and played internally, then unloaded and deleted.

### Usage

1. Select the object with the gradient.
2. Run the script.
3. Set the step count and the post-processing, then click OK.

### Options

**Steps**

Defaults to 5; any integer of 2 or more. Up/Down adjusts by 1, Shift+Up/Down snaps to multiples of 10.

**Post-processing**

- None: expand only
- Simple expand: expand the appearance afterwards
- Convert to blend: after expanding, build a blend from the two end objects

### Notes

- The action set and action names are kept in constants, and the names inside the .aia are generated from the same values.

### Article

[Split a gradient into solid-color objects (Japanese)](https://note.com/dtp_tranist/n/nbe084e691ba5)

### Update History

- v1.1.0 (2026-05-25)
