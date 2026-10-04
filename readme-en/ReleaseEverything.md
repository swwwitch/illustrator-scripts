# Release everything releasable, such as groups, compound paths, blends, and repeats, nested ones included

[![Direct](https://img.shields.io/badge/Direct%20Link-ReleaseEverything.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/ReleaseEverything.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseEverything.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Releases everything that can be released in the selection, such as groups, compound paths, compound shapes, blends, envelopes, repeats, and clipping masks, including nested ones.
- Clipping groups are released keeping both the mask path and the content, and the mask path gets a K100 fill at 15% opacity (no dialog).
- Only when a group contains other groups does a dialog let you choose "Release all levels" or "Release one level only".

### Main Features

- Releases the following repeatedly until none remain
  - Groups and compound paths
  - Compound shapes, blends, envelopes, Live Paint groups, and image tracing
  - Repeats (radial, grid, mirror) and Intertwine
- Releases clipping groups simply and applies a K100 fill at 15% opacity to the remaining mask path
- Also releases text wrap on the released objects
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
| Nested Groups | Release all levels | Releases everything nested (default) |
| | Release one level only | Releases the selected items once and keeps what is inside them |

The dialog appears only when a group contains other groups.

How clipping groups are released, and whether the mask path gets a fill, can be changed in `USER_DEFAULTS` at the top of the script.

| Setting | Values |
| --- | --- |
| `clipReleaseMode` | `"simple"` (keep the mask path and the content, default) / `"removePath"` (delete the mask path) / `"removeContent"` (delete the masked content) |
| `applyMaskFill` | `true` (default) applies the fill to the remaining mask path |

### Notes

- "Release all levels" also releases groups and compound paths inside the masked content. A remaining mask path that is a compound path stays a compound path.
- Compound shapes, blends, envelopes, and the like cannot be told apart by a script, so the Release commands are tried in turn and the script moves on with whichever one works.
- Each release gives the same result as the Release menu command: a blend leaves its spine path, an envelope leaves its envelope shape, Live Paint loses its fills, and image tracing returns to the original image.
- "Release all levels" also releases clipping groups inside groups, the same way as selected clipping groups (with the fill on the mask path).
- After the run, the objects that came out of the release are selected.

### Article

- [DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/nf5063dc9adae)

### Update History

- v1.0.0 (20261004) : Initial release
- v1.1.0 (20261004) : Blends, envelopes, Live Paint, image tracing, repeats, and Intertwine are now released too, along with text wrap. Renamed from ReleaseGroupsAndMasks to ReleaseEverything
- v1.1.1 (20261004) : Fixed clipping groups inside groups not being released
- v1.1.2 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
