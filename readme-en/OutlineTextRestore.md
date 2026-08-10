# OutlineTextRestore

[![Direct](https://img.shields.io/badge/Direct%20Link-OutlineTextRestore.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/outline/OutlineTextRestore.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OutlineTextRestore.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Restores outlined text back to live text using the information stored in the object's `note`.
- Only the selected paths and groups are processed, and the restored text is collected onto a `restored_text` layer.
- The original outlines (of the selection only) are moved to a newly created holding layer, which becomes `outlined_text`, is sent to the back, and is turned into a template layer.

### Main Features

- Extracts the text contents, font information and coordinates from the `note` of a PathItem or GroupItem
- Recreates the text frame at its original position and hides or moves the original outlines
- Warns when the note is missing or malformed
- Restored text always goes onto a new layer, and the original outlines always go onto a new holding layer (`outlined_text`, sent to the back and set as a template)

### Update History

- 20260111

### Script info

- Version: v1.3
