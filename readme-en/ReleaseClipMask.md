# Release clipping masks

[![Direct](https://img.shields.io/badge/Direct%20Link-ReleaseClipMask.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ReleaseClipMask.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseClipMask.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview:

- Illustrator script to release clipping masks based on selected modes.
- Supports simple release, remove mask path only, or remove masked image only.

<img alt="" src="https://www.dtp-transit.jp/images/ss-570-542-72-20250717-173004.png" width="50%" />

### Main Features:

- Simple release (keep both path and image)
- Remove mask path (keep placed image)
- Remove masked image (keep path)
- Optional K100 fill color (15% opacity) to path
- Japanese/English UI support
- Q/W/E hotkeys for mode selection

### Process Flow:

1. Choose mode and option in dialog
2. Release clipping mask accordingly
3. Optionally apply fill color to path

### Update History:

- v1.0 (20250606) : Initial release
- v1.1 (20250607) : Stabilization and adjustments
- v1.2 (20250717) : Comments refactored
- v1.2.3 (20260928) : The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.4 (20260928) : The button row is now built with the shared part
- v1.2.5 (20260929) : Dialog opacity changed to 98%
- v1.2.6 (20260930) : Fixed an error when running with characters selected by the Type tool
- v1.2.7 (20261001) : The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.2.8 (20261001) : Added space below the button row to match Illustrator's own dialogs
