# InspectKinsokuSimple

[![Direct](https://img.shields.io/badge/Direct%20Link-InspectKinsokuSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/InspectKinsokuSimple.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InspectKinsokuSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Collects the kinsoku set used by each paragraph in the selected text frames and lists the values in an alert.
- A small inspection script for checking whether kinsoku settings are mixed within a document.

### Usage

1. Select the text frames you want to inspect
2. Run the script
3. The detected kinsoku values are listed in an alert

### Notes

- Kinsoku "none" raises error 9563 because the attribute cannot be read, and is reported as "none".
- The document itself is never modified.
