# Add artboards that follow the existing layout

[![Direct](https://img.shields.io/badge/Direct%20Link-AddArtboardPlus.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddArtboardPlus.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddArtboardPlus.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Analyzes how the existing artboards are arranged in rows and columns and inserts new ones that follow the same pattern.

### Features

- Method: empty artboard, or duplicate of the current one (default from `DEFAULT_ADD_METHOD`)
- Position: after the current artboard, or at the end. Direction is right (horizontal) or down (vertical)
- Count: available only for empty artboards
- Spacing: estimated from the existing layout and shown in ruler units

### Usage

1. Make the reference artboard active.
2. Run the script.
3. Choose the method, position, count and spacing, then run it.

### Notes


### Update History

- v1.1.1
- v1.1.2 (20260914) : Fixed the layout direction or column count being misdetected from tiny coordinate differences, artboards being placed using the rounded display value when the spacing was left unchanged, and the gap after inserted artboards keeping the old spacing in "Added artboards only" mode, and values being rounded on keys other than Up/Down. Negative spacing is now treated as 0
- v1.1.3 (20260914) : Revised the UI wording (spacing scope is now "Apply to Added Only / Apply to All", and the alerts are more specific), added tooltips that explain each option and its shortcut key, and split the internal functions by role
- v1.2.0 (20260927) : Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.1 (20260928) : The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.2 (20260928) : The button row is now built with the shared part
- v1.2.2 (20260928) : Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held)
- v1.2.3 (20260929) : Dialog opacity changed to 98%
- v1.2.4 (20260930) : Fixed an error when running with characters selected by the Type tool
- v1.2.5 (2026-10-01) The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.2.6 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.2.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
