# Split a text frame into one frame per character

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartTextSplitter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SmartTextSplitter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartTextSplitter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Splits the selected text frame into one text frame per character, preserving the formatting.

### Usage

1. Select the text frame to split.
2. Run the script.

### Notes

- The split is always per character; there is no word or line mode.
- Use TextProcessingPalette.jsx to split by line or paragraph.

### Update History

- v2.0 (2026-02-17)
- v2.0.2 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v2.0.3 (2026-09-28): The button row is now built with the shared part
- v2.0.4 (2026-09-29): Dialog opacity changed to 98%
- v2.0.5 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v2.0.6 (2026-09-30): Button rows with only right-side buttons are now centered
- v2.0.7 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v2.0.8 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v2.0.9 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v2.0.10 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
