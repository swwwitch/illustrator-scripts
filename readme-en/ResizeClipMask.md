# ResizeClipMasks

[![Direct](https://img.shields.io/badge/Direct%20Link-ResizeClipMask.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ResizeClipMask.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResizeClipMask.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Automatically detects mask paths inside clip groups and replaces them with new masks adjusted by user-defined margins.
- Supports batch processing of multiple clip groups (rectangular masks only).

![](https://www.dtp-transit.jp/images/ss-404-262-72-20250710-044848.png)

### Main Features

- Detect and select mask paths
- Margin input via dialog
- Sign toggle button for margin inversion
- Rectangle detection and skip logic

### Workflow

1. Select clip groups
2. Specify margin in dialog
3. Detect rectangular masks and replace with new mask
4. Delete original mask path

### Change Log

- v1.0.0 (20250710): Initial release
- v1.3.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.3.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.2 (20260928): The button row is now built with the shared part. Mask detection now uses the shared part
- v1.3.3 (20260929): Dialog opacity changed to 98%
- v1.3.4 (20260930): Fixed an error when running with characters selected by the Type tool