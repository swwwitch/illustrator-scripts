# Select text frames by a combination of conditions

[![Direct](https://img.shields.io/badge/Direct%20Link-TextSelector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextSelector.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextSelector.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Selects text frames across the document by a combination of conditions.

### Features

- By attribute: search on font family, family plus style, family plus size, font size, text color or opacity, using the selected text as the reference
- By text kind: all, point text, area text or text on a path
- By string: exact, partial, prefix, suffix or regular expression
- Post-processing: none, hide, move to a `_text` layer, or bulk edit

### Usage

1. Select the reference text, if you are searching by attribute.
2. Run the script.
3. Set the conditions and the post-processing, then run it.

### Notes

- When moving to the `_text` layer, the lock and visibility state of existing layers is restored.
- Bulk edit replaces the contents while keeping the `characterAttributes` intact.

### Update History

- v1.2.5
- v1.2.7 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.8 (2026-09-28): The button row is now built with the shared part
- v1.2.8 (2026-09-28): Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held)
- v1.2.9 (2026-09-29): Dialog opacity changed to 98%
- v1.2.10 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.2.11 (2026-09-30): Button rows with only right-side buttons are now centered
- v1.2.12 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.2.13 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
