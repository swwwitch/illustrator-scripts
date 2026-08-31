# AddTrimMarkToCurrentArtboard

[![Direct](https://img.shields.io/badge/Direct%20Link-AddTrimMarkToCurrentArtboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AddTrimMarkToCurrentArtboard.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMarkToCurrentArtboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that always creates Japanese-style trim marks for the current artboard.
- It first gets or creates a dedicated トンボ (trim mark) layer and draws exactly one artboard rectangle on it.
- Trim marks are generated from that rectangle, and the same object is then turned into guides.
- If the トンボ layer is locked, it is unlocked for the run and restored to its original lock state afterwards.

### Main Features

- Always targets the current artboard
- Gets the トンボ layer automatically, creating it when missing
- Uses a single artboard rectangle (no duplicates)
- Generates the trim marks and the guides from the same object
- Preserves the original lock state of the トンボ layer

### Process Flow

1. Get the トンボ layer, creating it if it does not exist
2. Turn on Japanese-style trim marks in the preferences
3. Draw one artboard rectangle on the トンボ layer
4. Run the trim-mark command on that rectangle
5. Turn the same rectangle into guides
6. In `finally`, clear the selection and restore the トンボ layer's lock state

### note

- [【Illustrator】現在のアートボードにトンボを作成する｜DTP Transit 別館](https://note.com/dtp_tranist/n/n40e3e39cf9f2)

### Update History

- v1.0 (20250205): Initial version
- v1.1 (20260401): Unlock the トンボ layer for the run and restore its original lock state afterwards

### Script info

- Version: v1.1
