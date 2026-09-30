# Shift the baseline of specified characters

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartBaselineShifter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SmartBaselineShifter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBaselineShifter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Readme (GitHub):

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBaselineShifter.md

### Overview:

- Apply baseline shift individually to specified characters in selected text frames (in the East Asian Type unit set in Preferences > Units)
- Configure target characters and shift amount in dialog with instant preview

<img alt="" src="https://www.dtp-transit.jp/images/ss-742-402-72-20250716-205309.png" width="80%" />

### Main Features:

- Specify target characters
- Set shift amount as integer/decimal
- Reset all baseline shifts
- Instant preview and undo
- Japanese/English UI support

### Process Flow:

1. Select text frames
2. Configure in dialog
3. Check preview
4. Confirm with OK or revert with Cancel

### Update History:

- v1.0 (20240629): Initial version
- v1.3 (20240629): Added +/- buttons
- v1.4 (20240629): Two-column dialog layout, regex support
- v1.5 (20240630): Added function for TextRange selection
- v1.6 (20240630): Removed regex support, fine adjustments
- v1.7 (20250716): Refactoring, improved preview
- v1.8 (20250720): Added automatic calculation feature
- v2.2.2 (20260921): Shift amount now uses the unit set in Preferences, fixed decimals not being typable in the shift field, whitespace and line breaks left out of the default target, alert when auto calculation finds no target character
- v2.3.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v2.3.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v2.3.2 (20260929): Dialog opacity changed to 98%
- v2.3.3 (20260930): Fixed an error when running with characters selected by the Type tool
- v2.3.4 (20260930): Dropped the script's own rightward shift of the dialog on first open
- v2.3.5 (2026-10-01): Moved the buttons from a right-hand column to the standard bottom row (Reset on the left, Cancel/Adjust on the right). Unified the window and panel margins and spacing with the shared layout part
- v2.3.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
