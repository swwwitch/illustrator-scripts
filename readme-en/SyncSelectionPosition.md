# Align selected objects across documents to the same position

[![Direct](https://img.shields.io/badge/Direct%20Link-SyncSelectionPosition.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/SyncSelectionPosition.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SyncSelectionPosition.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Moves the selected objects in every other open document to the same position, using the top-left corner of the selection in the frontmost document as the reference.

### Usage

1. Open two or more documents.
2. Bring the reference document to the front and select an object in it.
3. Select the objects you want to align in the other documents as well.
4. Run the script.

### Notes

- If fewer than two documents are open, the script shows a warning and exits.
- It also exits with a warning when nothing is selected in the frontmost document.
- The reference point is the top-left (Left / Top) corner of the selection.
- Documents without a selection are skipped.
- Alert messages switch between Japanese and English according to the Illustrator UI locale.

### Article

- [Aligning selected objects across documents to the same position (Japanese)](https://note.com/dtp_tranist/n/n1f8155daeac4)

### Update History

- v1.0 (20251227): Initial version
- v1.0.1 (20260903): Reorganized the alert messages into LABELS with `getLabel()` and added Japanese/English switching
