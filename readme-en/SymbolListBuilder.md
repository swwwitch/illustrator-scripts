# Generate a "Symbol List" artboard

[![Direct](https://img.shields.io/badge/Direct%20Link-SymbolListBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/SymbolListBuilder.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolListBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Generates a dedicated "Symbol List" artboard that lays out every symbol registered in the Illustrator document.
Parameters are adjusted in a dialog with a live preview; OK commits the result (removes the preview, runs the final build, and cleans up the previous version).

### Main Features

- Placement reference: the last artboard, or a specified number (by default the artboard at the bottom-right of the canvas)
- Direction: build to the right of, or below, the reference artboard
- Size and padding: width, height and inner padding (the maximum width follows automatically when the width changes)
- Background: none / black / white / grey (K50); captions turn white on a black background
- Symbol filter: all symbols, or only those in use
- Captions: none / above / below, with a font size in Illustrator's text unit
- The default caption font follows the locale (ja: HiraginoSans-W3 / en: MyriadPro-Regular)
- Layer and artboard names follow the locale (Japanese "シンボル一覧" / English "Symbol List"), and existing ones are detected in either language
- With Update on, the existing Symbol List artboard and everything on it are deleted and replaced

### Units

- Ruler unit (`rulerType`) - sizes, margins and spacing
- Text unit (`text/units`) - font size

### Update History

- v1.0.0 (2026-05-09): Initial release
- v1.2.3 (2026-09-16): Code cleanup (shared helpers, split functions, fewer try blocks, clearer names). A corrupted saved setting now falls back to its default individually. The caption position default is unified to Bottom.
- v1.3.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.3.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.2 (2026-09-28): The button row is now built with the shared part. Settings are now saved through the shared part (stored in Folder.userData/illustrator-scripts/SymbolListBuilder.json)
- v1.3.3 (2026-09-29): Dialog opacity changed to 98%
- v1.3.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.3.5 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.3.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs

### Article

https://note.com/dtp_tranist/n/ncac687d0a3a0

### Script info

- Version: v1.3.3
