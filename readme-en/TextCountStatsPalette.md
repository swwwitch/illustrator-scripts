# Visualize statistics about the text in a document

[![Direct](https://img.shields.io/badge/Direct%20Link-TextCountStatsPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/TextCountStatsPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextCountStatsPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Visualize statistics of selected or all text objects in Illustrator
- Count characters, paragraphs, lines, English words, full-width characters, half-width kana, type (point/area/path), and fonts
- Persistent palette: keep it open, change selection, press Refresh to recount

### Main Features

- Shows a summary of various counts in a palette
- Automatically switches between selected objects and all content
- Delegates DOM counting to the main engine via BridgeTalk (avoids palette DOM disconnection)
- UI supports Japanese and English

### Process Flow

1. Show a persistent palette (existing one is closed first to prevent duplicates)
2. On Refresh (or right after showing), delegate counting to the main engine
3. Parse the marker-based result and update each panel value

### Update History

- v1.0 (20250806): Initial version
- v1.1 (20260702): Palette conversion (#targetengine + BridgeTalk delegation), refresh button, status line, localization cleanup
- v1.1.2 (2026-09-26) Renamed the file from `TextCountStats.jsx` to `TextCountStatsPalette.jsx`.
- v1.1.3 (2026-09-28) In Japanese, the status-line error text now uses a full-width colon.
- v1.1.3 (2026-09-28) Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held).
- v1.1.4 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.1.5 (2026-10-03) Moved to the shared persistent engine `SwwwitchPalettes` so CloseAllPalettes can close it
- v1.1.6 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

### Script info

- Version: v1.1.6
