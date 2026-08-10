# AiAnchorPointMarker

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAnchorPointMarker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/AiAnchorPointMarker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAnchorPointMarker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Places a marker (auto-generated square / frontmost object / symbol) at every anchor point of the selection, with a live preview you can tweak without closing the dialog.

- Choose what to add: an auto-generated square, a duplicate of the frontmost object, or an instance of a document symbol
- The square takes size (pt, decimals allowed, ⌘+↑↓ = ±0.1), color (a self-contained RGB dialog), and Symbolize (on by default; symbol named "アンカーポイント")
- Options: Scale (%, applied to the frontmost object / symbol), Move to layer (_anchorpoint), Group (on by default), and a 9-axis registration point
- In auto-generate mode, Scale and the 9-axis widget are dimmed and the registration point is fixed to center
- The live preview draws into a dedicated layer and is always cleaned up on OK / Cancel
- Edges and the Live Corner Annotator are hidden during the run (toggled on start and finish)
- Japanese / English localization; adapts to the light / dark UI

### Script info

- Version: v1.0.1
