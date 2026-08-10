# SmartAutoGroup

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAutoGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/SmartAutoGroup.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAutoGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script to automatically group selected objects based on conditions like "overlap", "vertical", "horizontal", or "proximity".
- Allows selecting mode and threshold via UI, and checks for ungrouped objects on retry.

### Main Features

- Switch modes: overlap, vertical, horizontal, proximity
- Threshold slider for distance (except overlap only)
- Auto-select grouped objects after grouping
- Prompt to retry if ungrouped objects remain
- Japanese and English UI support

### Process Flow

1. Configure mode and threshold in dialog
2. Extract groups using DFS traversal
3. Group objects using Illustrator's group command
4. Prompt to retry if any ungrouped objects remain

### Update History

- v1.0.0 (20250611): Initial version
