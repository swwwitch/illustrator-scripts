# Select every other object in stacking order

[![Direct](https://img.shields.io/badge/Direct%20Link-SelectAlternateItems.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SelectAlternateItems.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectAlternateItems.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Counts the selected objects in order and reselects only the odd- or even-numbered ones.

### Features

- Select switches between Odd and Even
- Count Order switches between Vertical, Horizontal and Stacking Order
- Preview inside the dialog
- Japanese / English UI

### Usage

1. Select several objects.
2. Run the script.
3. Choose Select and Count Order.
4. Confirm with OK.

### Notes

- If no document is open, or nothing is selected, the script shows a warning and exits.
- Vertical and Horizontal count by position; Stacking Order counts by stacking order.

### Update History

- v1.1.0
- v1.1.2 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.3 (2026-09-29): Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held)
- v1.1.4 (2026-09-29): Dialog opacity changed to 98%
- v1.1.5 (2026-09-29): Fixed the last Count Order not being saved. Renamed Direction to Count Order and centered the buttons. Stacking order now follows the selection order, so it counts correctly across groups and layers
