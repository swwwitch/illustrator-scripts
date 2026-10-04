# Toggle the video ruler, edges and other view items from a persistent palette

[![Direct](https://img.shields.io/badge/Direct%20Link-ViewTogglePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/ViewTogglePalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ViewTogglePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that toggles the video ruler, rulers in all documents, guides, smart guides, artboards, edges, canvas color and bounding box with one button each.

### Features

| Button | What it does |
| --- | --- |
| Video Ruler | Show/hide the video ruler |
| Show in All Documents | Turn the Show Rulers in All Documents preference (`useGlobalRulers`) on/off |
| Show Guides | Show/hide guides |
| Lock Guides | Lock/unlock guides |
| Smart Guides | Turn Smart Guides on/off |
| Artboards | Show/hide artboards |
| Edges | Show/hide edges |
| Canvas Color | Switch the canvas color between Match Brightness and White |
| Bounding Box | Show/hide the bounding box |

### Usage

Run the script and the palette opens. Each click on a button toggles that item. The palette also closes with `Esc` while it is active.

Running the script again while the palette is open brings the existing palette forward instead of opening a second one.

### Notes

A persistent palette cannot touch the DOM directly, so every toggle is delegated to the main engine over BridgeTalk.

The Canvas Color button rewrites `uiCanvasIsWhite` and then repaints the canvas with `zoomout` → `zoomin`.

The preferences you toggle most often live in [AiQuickPrefsPalette](AiQuickPrefsPalette.md).

### Update History

- v1.0.0 (2026-10-03) Split off from the View panel of AiQuickPrefsPalette
- v1.0.1 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
