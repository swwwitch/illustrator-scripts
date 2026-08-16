# RelinkSameImages.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-RelinkSameImages.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/RelinkSameImages.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkSameImages.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

Takes the selected placed image as a reference, finds every placed image in the document that references the same linked file, and relinks them all to a single file you choose.

When the same logo or shared asset sits on several artboards, the Links panel makes you relink them one at a time. This script needs one reference image and handles the rest in a single pass.

## Features

- Collects every placed image in the document that points at the same linked file as the selection
- Relinks all of them to the file you pick
- Resolves a placed image nested inside a group automatically, starting from the first selected item
- Matches on the absolute path, so a same-named file in another folder is left alone
- Skips embedded images and missing links
- Localized in Japanese and English

## Usage

1. Select one placed image. A group containing the image works too.
2. Run `RelinkSameImages.jsx`.
3. Pick the replacement file in the dialog.
4. Every placed image that referenced the same link is relinked, and the count is reported.

If several objects are selected, only the **first item in the selection** is used as the reference.

## How matching works

The script reads the reference image's link as an absolute path (`file.fsName`) and compares it against every placed image in the document (`document.placedItems`).

| Target | Relinked |
| --- | --- |
| Placed images referencing the same file | Yes |
| The reference image itself | Yes |
| Placed images inside groups and clip groups | Yes |
| A same-named file in another folder | No (different path) |
| Embedded images | No |
| Missing links | No |

To match on the file name only, or to select or delete the matches instead of relinking them, use `SelectSameLinks.jsx`.

## Notes

- There is no confirmation dialog. Use Undo (⌘Z / Ctrl+Z) to step back.
- Relinking to a different file format (PSD → PNG, for example) is allowed.
- The size of the result follows Illustrator's Preferences > File Handling setting for preserving dimensions on relink.
- If the reference image is embedded or its link is missing, the script reports it and stops.
- The whole document is processed: every artboard, and the inside of every group.
- The selection is cleared once the replacement finishes.
- Illustrator 2026 (June 2026 update) offers an equivalent bulk relink as a built-in feature.

## Article

[Relinking every placed image that shares a file name, in one step | DTP Transit](https://note.com/dtp_tranist/n/ne38eeee5abc8)

## Changelog

- v1.2.2 (2026-08-17): Tidied the header and metadata block, unified label lookup on `getLabel()`, added JSDoc, moved the body into `main()` (no behavior change)
- v1.2.1 (2026-06-18): Refactor — clearer naming, categorized labels, removed unnecessary try blocks and dead branches
- v1.2 (2026-03-09): Improved safety and maintainability
- v1.1 (2025-01-20): Improved resolution of placed images inside groups
- v1.0 (2024-06-15): Initial release
