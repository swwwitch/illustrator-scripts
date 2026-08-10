# CollectArtboardTexts

[![Direct](https://img.shields.io/badge/Direct%20Link-CollectArtboardTexts.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/CollectArtboardTexts.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CollectArtboardTexts.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Collects the text frames on every artboard and places them as new text frames, stacked vertically to the right of the last artboard.

### Main Features

- Picks the text frames overlapping each artboard
- Skips locked or hidden frames, and locked or hidden ancestor layers and groups
- Places one frame per source text, stacked vertically, rather than joining them with line breaks (a switch enables a combined mode)
- Selects the placed frames and zooms to fit

### Process Flow

1. Validate the active document and the target layer
2. Collect the texts contained in each artboard
3. Place each text as a separate frame, stacked vertically
4. Select the created frames and zoom in

### Update History

- v1.0.0 (20260513): Initial release

### Script info

- Version: v1.0.0
