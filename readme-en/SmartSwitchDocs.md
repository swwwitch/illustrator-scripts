# Switch quickly between open documents

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartSwitchDocs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/SmartSwitchDocs.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSwitchDocs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Quickly switches to another Illustrator document when several are open.

<img alt="" src="https://github.com/user-attachments/assets/e2e98c44-0db3-46c6-9f0e-579b17b82599" width="70%" />

### Main Features

- With **one** document open, nothing happens
- With **two**, switches to the inactive one without showing a dialog
- With **three or more**, pick the target from a list in a dialog (the document active at launch is left out of the list and shown under "Original Document")
- With Preview on (the default), the document switches as soon as it is selected in the list; with Preview off, it switches after OK is clicked
- Double-clicking a list item switches and closes, the same as OK
- Closing with Cancel, Esc or the close box returns to the original document
- Consolidates all document windows into tabs on launch (Consolidate All Windows)
- Japanese and English UI

### Usage

1. Run the script with two or more documents open
2. With three or more, pick the target with the ↑↓ keys or a click
3. Click OK or double-click an item to confirm, or Cancel to return to the original document

### Article

- [DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/nd9c7b7c077fb)

### Credits

- Continuous switching with the arrow keys and switching as soon as the dialog opens come from improvements by Yusuke Saegusa
  - [uske-s.hatenablog.com (Japanese)](https://uske-s.hatenablog.com/entry/2025/04/03/113115)

### Update History

- v1.0.0 (20250325): Initial version
- v1.1.0 (20250403): Returned focus to the dialog after arrow-key selection so documents can be switched in a row, switched to the first candidate when the dialog opens, and moved the layout to preferredSize (improvements by Yusuke Saegusa)
- v0.5.1 (20250525): Added Cancel button and adjusted UI
- v0.5.2 (20250525): Fixed focus maintenance after arrow key selection
- v0.5.3 (20260903): Added a Preview checkbox (switching waits for OK when it is off), put the version in the dialog title, fixed switching to the inactive document when exactly two are open; added the article URL to the basic info block, reorganized LABELS into nested categories with `getLabel()`, renamed variables/panels/functions to follow the naming rules, split dialog building and document collection into functions, unified duplicated selection/activation code, and added JSDoc to every function
- v0.5.4 (20260924): Closing the dialog with the window close box now also restores the original document, renamed the "Current Document" panel to "Original Document" and gave it a tooltip, and split part of the dialog building into functions
- v0.5.5 (20260924): Double-clicking a list item now switches and closes the dialog; added a tooltip to the list
