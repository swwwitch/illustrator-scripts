# ImageLinkManager

[![Direct](https://img.shields.io/badge/Direct%20Link-ImageLinkManager.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/ImageLinkManager.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageLinkManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Handles Embed, Unembed, Reset, Stroke and Relink for placed images (PlacedItem) from a single dialog.

The mode selector at the top switches the operation, and only the matching panel stays enabled.

### Features

- Embed: runs `embed()` on the selected, or all, placed items (PSD files are additionally ungrouped)
- Unembed: turns embedded images back into linked images
- Reset: resets the transformation of placed images
- Stroke: adds a stroke to placed images
- Relink: repoints the link
- Mode shortcuts: Embed **E** / Unembed **U** / Reset **R** / Stroke **S** / Link **L**

### Usage

1. Select the placed images (some modes can target the whole document).
2. Run the script.
3. Pick the mode at the top, set the panel options, and run it.

### Notes

- Use LinkedImageManager.jsx when you need full list-based management.

### Update History

- v1.2 (2025-12-21)
