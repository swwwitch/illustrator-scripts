# Merge the selection's bounds into one rectangle

[![Direct](https://img.shields.io/badge/Direct%20Link-BoundsToRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/BoundsToRectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BoundsToRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Merges the selected objects into a single rectangle based on their overall bounding box.

<img alt="The Merge Bounds into One Rectangle dialog" src="../png/ss-510-858-144-20260923-193029.png" width="30%" />

### Features

- Choose which object's fill and stroke to inherit (the key object, if set, is picked first)
- Switch between preview bounds and geometric bounds
- Option to keep the original objects
- Preview, with the center shown in the attributes panel

### Usage

1. Select the objects. To inherit a particular object's fill and stroke, click it once more to make it the key object.
2. Run the script.
3. Set the options and click OK.

### Article

https://note.com/dtp_tranist/n/nd4afdd8315f0

### Update History

- v1.4.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.4.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.4.5 (2026-10-01): The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.4.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.4.3 (2026-09-29): Dialog opacity changed to 98%
- v1.4.2 (2026-09-28): Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held)
- v1.4.2 (2026-09-28): The button row is now built with the shared part
- v1.4.2 (2026-09-28): Removed the space after the colon in English field labels, matching the Japanese labels. Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure
- v1.4.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.4.0 (2026-09-23): Added key object support (selected by default under "Take Fill and Stroke From"). Fixed the rectangle using geometric bounds when OK was clicked with "Use Preview Bounds" left on, and the rectangle sometimes landing on a layer other than the original active one. Revised the UI wording (title, panel name, tooltips), cleaned up the internals and added shortcut-key tooltips to the source radio buttons
- v1.3 (2026-03-08)
