# Generate a label with the font name and size

[![Direct](https://img.shields.io/badge/Direct%20Link-AddTextInfoLabel.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/AddTextInfoLabel.jsx)		

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTextInfoLabel.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script that displays font information below or to the right of selected text objects.
- You can switch between detailed display (area text) and compact display (point text), and customize which information to show.

### Main Features

- Displays detailed font info: font name, style, size, leading, kerning, tracking, etc.
- Toggle each display item on/off, plus buttons for All ON, All OFF, or Minimal Set
- Option to choose preview and position (bottom or right)
- Generated info text is placed on a "Font Info" layer
- ES3 compliant (Illustrator ExtendScript compatible)

### Process Flow

1. Choose display mode and items in the dialog
2. Check the preview
3. Press OK to generate and place the info text

### Update History

- v1.0.0 (20250420): Initial version
- v1.0.1 (20250423): Added item disable control and spacing settings
- v1.0.2 (20250424): Fixed incorrect leading (%) detection for accurate display
- v1.0.4 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.5 (20260928): Font size and leading are now converted to the units set in Preferences (the pt value used to be shown with the mm/Q label). Inches and picas now shown as "in" and "pica"
- v1.0.5 (20260928): The button row is now built with the shared part
- v1.0.6 (20260929): Dialog opacity changed to 98%
- v1.0.7 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.0.8 (20260930): Button rows with only right-side buttons are now centered
- v1.0.9 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones