# Reapply fill and stroke colors at random

[![Direct](https://img.shields.io/badge/Direct%20Link-ShuffleObjectColors.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ShuffleObjectColors.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShuffleObjectColors.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script to reapply fill and stroke colors randomly to selected objects (paths, text, groups, compound shapes).
- Offers options to exclude black and white, preserve color balance, and apply colors per character in text.

### Main Features

- Choose whether to apply to fill and/or stroke
- Options to exclude black and white
- Preserve color usage ratio (balance mode)
- Toggle between random or sequential application
- Random color application per character in text frames
- Japanese and English UI support

### Process Flow

1. Collect target objects (including groups and compound paths)
2. Collect colors based on specified conditions
3. Shuffle or order colors accordingly
4. Reapply colors to target objects

### Update History

- v1.0.0 (20240624): Initial release
- v1.0.1 (20240624): Bug fixes
- v1.0.2 (20240624): Added toggle for random/sequential application
- v1.0.3 (20240625): Localization adjustments
- v1.1.2 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.3 (20260929): Dialog opacity changed to 98%
- v1.1.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.1.5 (2026-10-01): Moved the buttons from a right-hand column to the standard bottom row (Apply on the left, Cancel/OK on the right). Unified the window and panel margins and spacing with the shared layout part
- v1.1.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.1.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)