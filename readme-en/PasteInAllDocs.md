# Paste into every open document at the same position

[![Direct](https://img.shields.io/badge/Direct%20Link-PasteInAllDocs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/PasteInAllDocs.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PasteInAllDocs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Pastes the copied objects into every open document at the same position.

It uses Paste in Place, so the objects land on the coordinates of each document.

### Usage

1. Copy the objects you want to paste.
2. With all the destination documents open, run the script.

### Notes

- If nothing has been copied, the script shows a warning and exits.
- The source document does not receive a duplicate.
- Each destination document is switched to the same artboard index as the source before pasting.
- If a destination document's active layer is locked or hidden, it is unlocked for the paste and restored afterwards.
- Documents that could not receive the objects are listed together at the end.
- Selections are cleared in every document the script touches.
- Messages appear in Japanese or English, depending on the application locale.

### Article

[Paste into every open document with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n04535658c7f6)

### Update History

- v1.1.0 (2026-09-19): Localized the messages (Japanese/English), fixed the duplicate left in the source document, the offset caused by a mismatched active artboard and the undetected paste failures, and added a link to the article
- v1.0
