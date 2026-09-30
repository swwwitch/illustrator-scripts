# Select placed images that share the same link

[![Direct](https://img.shields.io/badge/Direct%20Link-SelectSameLinks.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/SelectSameLinks.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameLinks.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Finds every PlacedItem in the active document that references the same linked file (or files) as the current selection, then selects or deletes them according to the dialog options.

### Dialog

- Matching: same path (absolute path) / same filename (path ignored)
- Action: select the matching links / delete the linked images only / delete them together with their clip groups

### Update History

- v1.1.8 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.1.7 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.1.6 (2026-09-30): Button rows with only right-side buttons are now centered
- v1.1.5 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.1.4 (2026-09-29): Dialog opacity changed to 98%
- v1.1.3 (2026-09-28): The button row is now built with the shared part
- v1.1.2 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%

### Script info

- Version: v1.1.6
