# RandomizeObjects

[![Direct](https://img.shields.io/badge/Direct%20Link-RandomizeObjects.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RandomizeObjects.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RandomizeObjects.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Randomly moves, scales, rotates and changes the opacity of the selected objects
- Parameters are set in the dialog and the result is confirmed with an immediate preview
- Colour supports both a normal shuffle and a full shuffle
- Reset restores the state from before the dialog opened
- The "avoid overlap" placement logic is factored into a function that other scripts can reuse

### Main Features

- Move distance set separately for the horizontal and vertical axes, or linked
- A center option gathers the objects in one place
- Randomized scaling, rotation and opacity
- Preview applied live, with a full reset to the original state on cancel
- More stable UI state syncing (checkbox state and field enablement stay consistent)
- The preview is recomputed from the base state every time, so effects no longer accumulate

### Process Flow

1. Check the document and the selection
2. Show the dialog (move distance, scale ratio, rotation, opacity)
3. Apply the randomized transform with a live preview
4. Commit on OK, or restore the original state on cancel

### Update History

- v2.2 (2026-03-05): Extracted the "avoid overlap" placement logic into a general-purpose function
- v2.1 (2026-02-27): Full shuffle now generates random colours (K is 0-30 in CMYK); colour runs under Random

### Script info

- Version: v2.2
