# CreateGuidesFromSelection

[![Direct](https://img.shields.io/badge/Direct%20Link-CreateGuidesFromSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/CreateGuidesFromSelection.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script to create guides from selected objects in Illustrator.
- Flexible dialog UI to specify top, bottom, left, right, and center guides.

![](https://www.dtp-transit.jp/images/ss-756-962-72-20250712-220853.png)

### Main Features

- Target the "Artboard" or the pseudo "Canvas"
- Extend guides beyond the artboard, or offset them from the target
- Use preview or geometric bounds
- Temporary outlining and appearance expansion for text objects
- Destination: the same layer as the selection, or the "_guide" layer
- "_guide" layer management and guide removal option
- Live preview (existing guides are hidden while it is shown)
- Supports clip groups and multi-selection

### Workflow

- Configure options in dialog (the live preview follows every change)
- Get the bounding box of the target (selection / artboard / canvas)
- Draw guides
- Temporarily outline and expand appearance for text, then restore
- Lock the "_guide" layer when the guides are drawn there

### Update History

- v1.0 (20250711): Initial version
- v1.1 (20250711): Multi-selection & clip group support, offset & bleed features, text outline support
- v1.2 (20250711): Added appearance expansion, UI improvements, enhanced error handling
- v1.3 (20250712): Code refactor and radio button visibility toggle
- v1.6 (20250712): refactored code, added radio button visibility toggle feature
- v1.6.1 (20250712): Minor adjustments
- v1.6.2 (20250712): Unit settings
- v1.8.0 (20250802): UI cleanup, wording and tooltip adjustments
- v1.9.0 (20260628): Live preview, structured localization, existing guides hidden during preview
- v1.9.1 (20260803): Added the "Destination" panel (selection layer / "_guide" layer), added "Group the guides to draw", reworded the per-object option, dimmed the margin field for center-only guides, fixed bounds for mixed text and object selections, internal refactoring