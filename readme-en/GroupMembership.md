# Add to, release from, or group each, chosen in a dialog

[![Direct](https://img.shields.io/badge/Direct%20Link-GroupMembership.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/GroupMembership.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupMembership.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Choose in a dialog whether to add the selection to an existing group, release it from its groups, or wrap each object in its own group.
- The default is picked from the selection: "Release from group" when objects inside a group are selected, "Add to existing group" for one group plus other objects, and "Group each object" otherwise.
- One-click versions without a dialog are also available ([AddToGroup](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddToGroup.md) / [ReleaseFromGroup](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseFromGroup.md) / [GroupEachSelection](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEachSelection.md)).

### Main Features

- Add to existing group: moves the other selected objects into the group selected with them. The group is not released, so its opacity, effects, name and clipping mask are kept, and the stacking order is preserved. With no group or several groups, the groups are released one level and everything is regrouped as one (clipping groups are kept intact).
- Release from group: moves the objects selected inside groups out of their groups.
- Group each object: wraps each selected object in its own group without changing the stacking order.

### Usage

1. Select the objects (to release objects from a group, select them inside the group with the Direct Selection tool or similar)
2. Run the script, choose the action and click OK

### Options

- Place at (Release from group): "Top of the layer" or "Just in front of the group". The latter keeps the objects in place relative to other objects.
- Skip objects that are already groups (Group each object): selected groups are left as they are.

### Notes

- "Add to existing group" needs two or more selected objects, and "Release from group" needs objects selected inside a group.
- The script does not run while text is selected with the Type tool.

### Article

- [DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n36fbd4162721)

### Change Log

- v1.0.0 (20260922): Initial release
