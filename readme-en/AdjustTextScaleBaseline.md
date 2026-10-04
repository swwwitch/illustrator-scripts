# Adjust Text Scale Baseline

[![Direct](https://img.shields.io/badge/Direct%20Link-AdjustTextScaleBaseline.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AdjustTextScaleBaseline.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustTextScaleBaseline.md)
)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description:

- Adjust font size, horizontal/vertical scale, baseline shift, kerning, and tracking of selected text in Illustrator
- Realtime preview and numeric input via dialog box

<img alt="" src="https://www.dtp-transit.jp/images/ss-782-654-72-20250724-092616.png" width="70%" />

### Main Features:

- Convert between apparent size and font size
- Individually adjust horizontal/vertical scale
- Modify baseline shift, kerning, and tracking
- Confirm with OK, restore with Reset

### Processing Flow:

1. Build UI and define labels (ja/en)
2. Get initial state of selected text
3. Apply changes in realtime on input
4. Confirm with OK, cancel or reset to revert

### Change Log:

- v1.0 (20250720): Initial release
- v1.1 (20250724): Added kerning, baseline shift, and tracking; refactored UI and event logic
- v1.2 (20250725): Enhanced preview handling, dim apparent size at 100% scale, fixed shift key increments
- v1.3 (20250726): Added tracking feature, UI adjustments
- v1.5.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.5.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.5.2 (20260929): Dialog opacity changed to 98%
- v1.5.3 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.5.4 (20260930): Dropped the script's own rightward shift of the dialog on first open
- v1.5.5 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part

### Update History

- v1.5.6 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
