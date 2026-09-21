# Ungroup nested subgroups recursively

[![Direct](https://img.shields.io/badge/Direct%20Link-SimplifyGroups.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/SimplifyGroups.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SimplifyGroups.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Recursively ungroups subgroups inside the selected group.
- The outermost group remains intact.
- Other selected objects are moved into the group. With several groups, everything is grouped together first.

### Main Features

- Recursive ungrouping of subgroups
- Moves other selected objects into the group (auto-groups when there are several groups)
- Uses Illustrator's "ungroup" menu command internally

### Process Flow

1. Check document and selection
2. If other objects are selected, move them into the group, or group everything when there are several groups
3. Recursively find and ungroup subgroups

### Notes

- Locked or hidden subgroups are left as they are.

### Change Log

- v1.0.0 (20250707): Initial release
- v1.3.1 (20260922): Reworked the processing; objects can now be moved into clipping groups masked by text; locked or hidden subgroups are now left intact