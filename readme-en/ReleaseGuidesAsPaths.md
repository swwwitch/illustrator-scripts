# Release guides on the _guide layer as paths

[![Direct](https://img.shields.io/badge/Direct%20Link-ReleaseGuidesAsPaths.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/ReleaseGuidesAsPaths.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseGuidesAsPaths.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

This script targets a layer named `_guide` in the active Adobe Illustrator document.

- Unlocks the target layer
- Removes the guide attribute from every item on that layer and gives it no fill and a 1pt K100 stroke
- Moves those items to a layer named `ReleasedGuides` (created if missing; an existing `UnlockedGuides` layer from earlier versions is reused instead)
- Locks the `_guide` layer again afterwards

### Update History

- v1.0.1 (20260916): Renamed from unlockGuideLayerAndClearGuides to ReleaseGuidesAsPaths. Fixed the stroke not becoming K100. Shows an alert when no document is open or no "_guide" layer exists. The destination layer is now "ReleasedGuides" (an existing "UnlockedGuides" layer is still reused)
- v1.0 (20250716): Initial version

### Script info

- Version: v1.0.1
