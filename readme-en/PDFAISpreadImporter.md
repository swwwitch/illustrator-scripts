# Import PDF/AI pages as artboards, splitting spreads (merged into PDFAIImporter)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

This script has been merged into [PDFAIImporter.jsx](PDFAIImporter.md).

To split spreads, turn on Split left and right in PDFAIImporter's Spreads panel. Placing into a new document, choosing the crop box and the even-page side (detected from the PDF binding direction) are all available in PDFAIImporter.

### Article

[Import a spread PDF one page at a time in Illustrator (Japanese)](https://note.com/dtp_tranist/n/n5514d9f2c5f8)

### Update History

- 2026-09-29: Merged into PDFAIImporter.jsx and removed
- v1.1.4 (2026-09-28): The button row is now built with the shared part
- v1.1.3 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.2 (2026-09-25): Added colons to field labels, a label and tooltip for the crop box and color mode, renamed the button to “Choose File...” and the panel to “Pages”, fixed the crop box values so Trim, Bleed and Art place the box you choose, wrapped artboards onto a new row at the canvas edge to fix an error with many pages, and removed internal duplication
- v1.1.1 (2026-09-19): Added a link to the article and tooltips to each option, and reorganised the internal naming and structure
- v1.1.0 (2026-03-18)
