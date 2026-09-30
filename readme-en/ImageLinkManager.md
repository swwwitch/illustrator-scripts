# Embed, unembed and relink placed images in one dialog

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

- Use LinkedImageManagerPalette.jsx when you need full list-based management.

### Update History

- v1.3.4 (2026-09-30) : Fixed an error when running with characters selected by the Type tool
- v1.3.3 (2026-09-29) : Dialog opacity changed to 98%
- v1.3.2 (2026-09-28) : Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held)
- v1.3.2 (2026-09-28) : The button row is now built with the shared part
- v1.3.1 (2026-09-28) : The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.0 (2026-09-27) : Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2 (2025-12-21)
