# SymbolListBuilder

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

### Article

https://note.com/dtp_tranist/n/ncac687d0a3a0

### Script info

- Version: v1.2.2
