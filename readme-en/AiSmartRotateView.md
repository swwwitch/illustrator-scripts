# AiSmartRotateView

[![Direct](https://img.shields.io/badge/Direct%20Link-AiSmartRotateView.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiSmartRotateView.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSmartRotateView.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Get the active view rotation angle (view rotation) of the current document
- Show the rotation angle in a palette
- Click "Apply" to set the "constrain angle" preference (Shift-key angle)

With "Link to view rotation" on, working with the view rotated 30 degrees seeds the constrain angle
with the same 30 degrees, so the view and operation angles match for easier work.

Because this is a palette (persistent window), DOM access (reading the view rotation and
applying the preference) is delegated to the main engine via BridgeTalk on each button press.

### Script info

- Version: v1.0.0
- First release: 20260605
- Last updated: 20260805
