# ApplyLeadingPerTextFrame

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyLeadingPerTextFrame.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ApplyLeadingPerTextFrame.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyLeadingPerTextFrame.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Recalculates the leading of each line in the selected text frames from the font size of the first few characters, and applies it.

The leading percentage is chosen in a dialog.

### Usage

1. Select the text frames.
2. Run the script.
3. Set the leading percentage and run it.

### Notes

- A partial selection (a TextRange) is normalized to its parent text frame.
- The adjustment changes the auto-leading percentage rather than the leading value itself.
- ApplyLeadingPerTextFrame110.jsx, 150.jsx and AUTO.jsx are fixed-percentage variants.

### Update History

- v1.1.0 (2026-07-08)
