# Release groups, compound paths, compound shapes, and clipping masks, nested ones included

[![Direct](https://img.shields.io/badge/Direct%20Link-ReleaseGroupsAndMasks.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/ReleaseGroupsAndMasks.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseGroupsAndMasks.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Releases the groups, compound paths, compound shapes, and clipping masks in the selection, including nested ones.
- Clipping groups are released keeping both the mask path and the content, and the mask path gets a K100 fill at 15% opacity (no dialog).
- Only when a group contains other groups does a dialog let you choose "Release all levels" or "Release one level only".

### Main Features

- Releases groups, compound paths, and compound shapes repeatedly until none remain
- Releases clipping groups simply and applies a K100 fill at 15% opacity to the remaining mask path
- Nested groups: "Release all levels" or "Release one level only"
- Mixed selections are released one object at a time
- Japanese / English UI

### How to Use

1. Select the objects to release
2. Run the script
3. If a dialog appears, choose how deep to release and click OK

### Options

| Panel | Option | Description |
| --- | --- | --- |
| Nested Groups | Release all levels | Releases every nested group, compound path, and compound shape (default) |
| | Release one level only | Releases the selected items once and keeps what is inside them |

The dialog appears only when a group contains other groups.

How clipping groups are released, and whether the mask path gets a fill, can be changed in `USER_DEFAULTS` at the top of the script.

| Setting | Values |
| --- | --- |
| `clipReleaseMode` | `"simple"` (keep the mask path and the content, default) / `"removePath"` (delete the mask path) / `"removeContent"` (delete the masked content) |
| `applyMaskFill` | `true` (default) applies the fill to the remaining mask path |

### Notes

- "Release all levels" also releases groups and compound paths inside the masked content. A remaining mask path that is a compound path stays a compound path.
- Compound shapes are released with the Release Compound Shape action. Blends and envelopes are left as they are.
- Nested clipping groups are released as plain groups, not as masks (the mask path remains with no fill or stroke).
- After the run, the objects that came out of the release are selected.

### Update History

- v1.0.0 (20261004) : Initial release
