# CreateGuidesFromSelection

[![Direct](https://img.shields.io/badge/Direct%20Link-CreateGuidesFromSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/CreateGuidesFromSelection.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script to create guides from selected objects in Illustrator.
- Flexible dialog UI to specify top, bottom, left, right, and center guides.

![](https://www.dtp-transit.jp/images/ss-756-962-72-20250712-220853.png)

### Main Features

- Use preview or geometric bounds
- Temporary outlining and appearance expansion for text objects
- Offset and bleed (margin) support
- "_guide" layer management and guide removal option
- Supports clip groups and multi-selection

### Workflow

- Configure options in dialog
- Get bounding box of selected objects
- Draw guides
- Temporarily outline and expand appearance for text, then restore
- Add guides to "_guide" layer and lock

### Update History

- v1.0 (20250711): Initial version
- v1.1 (20250711): Multi-selection & clip group support, offset & bleed features, text outline support
- v1.2 (20250711): Added appearance expansion, UI improvements, enhanced error handling
- v1.3 (20250712): Code refactor and radio button visibility toggle
- v1.6 (20250712): refactored code, added radio button visibility toggle feature
- v1.6.1 (20250712): Minor adjustments
- v1.6.2 (20250712): Unit settings