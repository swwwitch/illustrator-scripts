# SmartAutoGroup-yoko

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAutoGroup--yoko.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/SmartAutoGroup-yoko.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAutoGroup-yoko.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Automatically groups the selection into horizontal rows.

Objects count as the same row when their vertical offset is within the tolerance, however far apart they are horizontally.

### Features

- Horizontal grouping (detects rows by vertical alignment; horizontal distance is not considered)
- A Vertical tolerance slider for tuning the row detection
- Per-artboard option, so objects on different artboards never share a group
- The created groups are selected once grouping finishes
- A prompt to re-run when objects were left ungrouped
- Japanese / English UI

### Usage

1. Select all the objects you want to group.
2. Run the script.
3. Adjust Vertical tolerance and click Group.

### Options

**Vertical tolerance**

Two objects belong to the same row when their vertical offset — the gap between them, when they do not overlap — is within this value. The unit is points (default 10 pt, range 0–200 pt). Horizontal distance is never considered, so objects at opposite edges of the artboard still form one row as long as they line up vertically. At 0, only objects that actually overlap vertically are grouped. The test uses `visibleBounds`.

**Per artboard**

When on, objects on different artboards end up in separate groups even if they line up vertically. An object belongs to whichever artboard rectangle contains its center point. The initial state depends on the situation: dimmed and off when there is only one artboard, on when the selection spans several artboards, and off otherwise.

### Notes

- Rows are built by walking from neighbour to neighbour (connected components via DFS). If A and B are the same row and B and C are the same row, A and C end up in one group even when their own vertical offset exceeds the tolerance. Lower the tolerance if objects chain together unexpectedly.
- A row with only one member is not grouped. When such objects remain, their count is reported and you can reopen the dialog to try a wider tolerance.
- A progress bar is shown while grouping.
- Grouping uses Illustrator's standard group command, which moves the new group to the frontmost layer; the script moves it back to the layer the objects came from.

### Update History

- v1.0.0 (2025-06-11): First release
- v1.1.0 (2026-06-09): Specialized to horizontal grouping (mode selection removed). Row detection now uses the Vertical tolerance threshold, with no limit on horizontal distance
- v1.2.0 (2026-06-09): Added the Per artboard checkbox. Defaults to on only when the selection spans several artboards, and is dimmed when there is a single artboard. Added a progress bar, and moved the slider onto the line below its label
