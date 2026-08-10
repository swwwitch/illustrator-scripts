# PathInspector

[![Direct](https://img.shields.io/badge/Direct%20Link-PathInspector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/PathInspector.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathInspector.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Count path stats for the selection / whole document and show them in a persistent palette
- Export the stats as a plain text report
- Persistent palette: keep it open, change selection, press Refresh to recount
- Delegate DOM counting to the main engine via BridgeTalk

### Main Features

- Path stats: open/closed, anchor points, handles, compound paths, compound shapes
  * Guide paths (guides=true) are excluded from path stats
- Recursively count paths inside groups
- Export button writes a plain text report to the desktop
- Remember and restore the palette position during the current session

### Process Flow

- Show a persistent palette (existing one is closed first to prevent duplicates)
- On Refresh (or right after showing), delegate counting to the main engine
- Parse the marker-based result and update the path panel
- Export is generated from the collected data

### Script info

- Version: v1.0.0
- First release: 20260731
- Last updated: 20260731
