# FillSnapper

[![Direct](https://img.shields.io/badge/Direct%20Link-FillSnapper.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/FillSnapper.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillSnapper.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Classifies the current selection into items to move and snap references, then snaps each item's bounding box to the nearest reference line.

Paths are transformed at the anchor level, so child paths inside clipping groups are handled too.

### Usage

1. Select both the objects to move and the rules to snap to.
2. Run the script.
3. Set the tolerance and the snap distance, then run it.

### Options

**Line detection tolerance**

How thin a path has to be before it is treated as a horizontal or vertical line.

**Maximum snap distance**

How far from an edge a reference line may sit and still attract it. 0 means no distance limit.

### Update History

- v1.0.1
