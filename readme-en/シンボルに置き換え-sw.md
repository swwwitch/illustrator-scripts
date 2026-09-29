# Replace the selection with a symbol chosen from the document

[![Direct](https://img.shields.io/badge/Direct%20Link-%E3%82%B7%E3%83%B3%E3%83%9C%E3%83%AB%E3%81%AB%E7%BD%AE%E3%81%8D%E6%8F%9B%E3%81%88--sw.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/%E3%82%B7%E3%83%B3%E3%83%9C%E3%83%AB%E3%81%AB%E7%BD%AE%E3%81%8D%E6%8F%9B%E3%81%88-sw.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/%E3%82%B7%E3%83%B3%E3%83%9C%E3%83%AB%E3%81%AB%E7%BD%AE%E3%81%8D%E6%8F%9B%E3%81%88-sw.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Picks a symbol from the ones registered in the document and replaces the selected objects with instances of it.

### Usage

1. Select the objects you want to replace.
2. Run the script.
3. Choose a symbol from the list and click OK.

### Notes

- The list shows a limited number of symbols at a time; scroll to reach the rest.
- Large selections can take a while to process.
- Based on "シンボルに置き換え.jsx" by Toshiyuki Takahashi ([graphicartsunit.com](http://www.graphicartsunit.com/)).

### Update History

- v0.5.0
- v0.5.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v0.5.2 (2026-09-28): Removed the trailing space after the colon in English field labels ("Symbol: " -> "Symbol:"). The button row is now built with the shared part. Fixed the dialog failing to open because it was shown before LABELS were set, and the title showing version 0.5.0
- v0.5.3 (2026-09-29): Dialog opacity changed to 98%
- v0.5.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
