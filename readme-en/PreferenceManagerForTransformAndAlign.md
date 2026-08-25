# PreferenceManagerForTransformAndAlign

[![Direct](https://img.shields.io/badge/Direct%20Link-PreferenceManagerForTransformAndAlign.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PreferenceManagerForTransformAndAlign.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManagerForTransformAndAlign.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette for switching a range of Illustrator preferences. Every change takes effect the moment you make it.

### Features

- Runs as a palette on a resident engine; writes are delegated to the main engine via BridgeTalk, while reads happen synchronously in place
- Two columns: keyboard increment and transform/align on the left, align-to-glyph-bounds and guides/rulers on the right, with artboard names and borders plus other settings across the bottom
- Changes made elsewhere — in the Preferences dialog, for example — are picked up when the palette is clicked back into focus

### Usage

1. Run the script to open the palette.
2. Toggle the settings; each one applies immediately.

### Notes

- The palette is a `#targetengine` resident script: after editing the code, **close the palette before running it again**, otherwise the old code keeps running.

### Update History

- v1.6.0 (2026-06-27)
