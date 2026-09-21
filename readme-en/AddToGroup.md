# Add objects to an existing group

[![Direct](https://img.shields.io/badge/Direct%20Link-AddToGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/single-function/AddToGroup.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddToGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Adds the selected objects to the existing group selected with them.
- The group is not released, so its opacity, effects, name and clipping mask are kept. Objects added to a clipping group are masked.
- The stacking order is preserved: objects in front of the group go to the top of the group, objects behind it go to the bottom.
- With no group or several groups, the groups are released one level and everything is regrouped as one. Clipping groups are kept intact.

### Usage

1. Select the objects and the group you want to merge them into
2. Run the script

### Notes

- Does nothing when fewer than two objects are selected or when text is selected with the Type tool.
- There is no dialog.

### Article

- [DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n36fbd4162721)

### Change Log

- v1.0.0 (20260306): Initial release
- v1.0.2 (20260922): Objects are now added without releasing the group, so clipping masks and the group's effects and name are kept. Does nothing while text is selected with the Type tool
