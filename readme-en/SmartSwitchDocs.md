# SmartSwitchDocs

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartSwitchDocs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/SmartSwitchDocs.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script to quickly switch to another Illustrator document when multiple documents are open.
- Automatically switches if only two documents are open; shows a dialog for selection when there are three or more.

### Main Features

- Auto switch depending on the number of open documents
- Select from a list when there are three or more documents
- Immediate preview switching inside the dialog
- Revert to original document on cancel
- Japanese and English UI support

### Process Flow

1. Check the number of open documents
2. If two, automatically switch to the inactive document
3. If three or more, show dialog and select from the list
4. Switch after pressing OK or Cancel

### Update History

- v1.0.0 (20250325): Initial version
- v0.5.1 (20250525): Added Cancel button and adjusted UI
- v0.5.2 (20250525): Fixed focus maintenance after arrow key selection
