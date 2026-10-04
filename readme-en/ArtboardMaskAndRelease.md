# Mask each artboard with a rectangle

[![Direct](https://img.shields.io/badge/Direct%20Link-ArtboardMaskAndRelease.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ArtboardMaskAndRelease.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardMaskAndRelease.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Draws the same-sized rectangle on all artboards and masks objects inside each artboard.
- Sets the clip group name to the artboard name.
- Duplicates objects that span multiple artboards and masks them on each.
- Includes a release function with an option to ungroup after release.
- Allows margin adjustment to control the mask area.

### Main Features

- Apply mask per artboard
- Set clip group name to artboard name
- Duplicate objects spanning multiple artboards
- Release mask and optional ungroup
- Margin setting
- Japanese / English UI support

### Workflow

1. Select mode (Mask / Release) and options in the dialog
2. When "Mask" is selected, create rectangles for each artboard and mask objects
3. When "Release" is selected, release clipping and optionally ungroup

### Update History

- v1.0 (20250710) : Initial version
- v1.1 (20250710) : Added option to remove objects outside artboards, options for locked/hidden objects
- v1.2.0 (20260927) : Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.1 (20260928) : The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.2 (20260928) : The button row is now built with the shared part
- v1.2.3 (20260929) : Dialog opacity changed to 98%
- v1.2.4 (20260930) : Fixed an error when running with characters selected by the Type tool
- v1.2.5 (20260930) : Button rows with only right-side buttons are now centered
- v1.2.6 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.2.7 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.2.8 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.2.9 (2026-10-03) The number field now shows its unit inside the field; removed the unit label to the right of the field
- v1.2.10 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

### Script info

- Version: v1.2.10
