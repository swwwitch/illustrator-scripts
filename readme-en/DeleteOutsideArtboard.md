# Delete Objects Outside Artboards

[![Direct](https://img.shields.io/badge/Direct%20Link-DeleteOutsideArtboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/DeleteOutsideArtboard.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeleteOutsideArtboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- This script checks objects in the document against artboards and deletes or moves objects outside to a backup layer.
- You can select "Current Artboard Only" or "All Artboards" as target.

![](https://www.dtp-transit.jp/images/ss-428-504-72-20250708-023849.png)

### Main Features

- Delete objects outside artboards
- Option to move to backup layer
- Japanese/English interface support

### Workflow

1. Select target artboards and move option in dialog
2. Check object overlap with artboards
3. Remove or move non-overlapping objects

### Changelog

- v1.0.0 (20250708): Initial version
- v1.4.2 (20260927): Fixed Outside Artboard: Delete doing nothing when combined with Move to Backup Layer. Added a no-document alert and refined the tooltips
- v1.4.3 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.4.4 (20260928): The button row is now built with the shared part
- v1.4.5 (20260929): Dialog opacity changed to 98%
- v1.4.6 (20260930): Fixed an error when running with characters selected by the Type tool
