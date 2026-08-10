# PathCleanupTool-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-PathCleanupTool-v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/PathCleanupTool-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathCleanupTool-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Updated: 2026-03-20
- Optimizes the selected paths (including those inside groups and compound paths) by removing redundant anchors, duplicate anchors and Bezier handles that can be treated as straight
- The Other tab also offers smoothing, cornering, adding anchors and splitting at anchors

### Main Features

- Removes anchor points that are redundant along a straight run
- Removes anchor points that share the same coordinates
- Removes handles on Bezier segments that can be treated as straight
- Locked and hidden objects (including their parents and layers) are skipped automatically
- The selection is frozen when the dialog opens, so the information shown matches what is processed
- Separate tolerances for anchor removal and handle removal
- Smoothing guards against neighbouring anchors that coincide or sit extremely close
- Open-path endpoints use a natural tangent direction rather than wrapping around

### Script info

- Version: v1.4.1
