# SimplifyGroups.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-SimplifyGroups.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/SimplifyGroups.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Recursively ungroups subgroups inside the selected group.
- The outermost group remains intact.
- If groups and non-groups are mixed in the selection, they are grouped together before processing.

### Main Features

- Recursive ungrouping of subgroups
- Auto-grouping when mixed selection
- Uses Illustrator's "ungroup" menu command internally

### Process Flow

1. Check document and selection
2. Group mixed selection if needed
3. Recursively find and ungroup subgroups

### Change Log

- v1.0.0 (20250707): Initial release