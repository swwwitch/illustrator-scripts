# OverlapRemover.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-OverlapRemover.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/OverlapRemover.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OverlapRemover.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script runs a sequence of menu commands against the current selection to flatten overlaps.

The order is Offset Path, Group, Pathfinder: Merge, then Expand Appearance. The objects produced by Offset Path are reselected together with the original selection, so the following commands act on the intended set.

## Main features

- Operates only on the current selection
- Detects newly created objects by comparing the document before and after Offset Path
- Hands the union of the original selection and the new objects to the following commands
- Skips locked and hidden objects when building the selection
- Aborts with an alert when a command fails
- Japanese and English UI

## Usage

1. Select the objects whose overlaps you want to flatten.
2. Run `OverlapRemover.jsx`.
3. The Offset Path dialog opens. Enter the offset value and click OK.
4. Group, Pathfinder: Merge, and Expand Appearance then run automatically.

## Command sequence

| Step | Action | Menu command |
| --- | --- | --- |
| 1 | Offset Path | `OffsetPath v22` |
| 2 | Group | `group` |
| 3 | Pathfinder: Merge | `Live Pathfinder Merge` |
| 4 | Expand Appearance | `expandStyle` |

## Notes

- Offset Path opens a dialog, so the offset value is entered by the user.
- The script stops with a message when no document is open or nothing is selected.
- Locked and hidden objects are never included in the selection.
- Illustrator has no undo-grouping API, so undoing this run takes multiple steps. A single Undo will not revert everything.

## Changelog

- v1.0.0 (20260301): Initial version
- v1.0.1 (20260727): Dropped the "single Undo step via suspendHistory" claim (`suspendHistory` is a Photoshop API and is never called under Illustrator). Also added the basic-info and localization sections, replaced the Japanese-only messages with bilingual ones, and folded the repeated selectability checks into `isSelectable`
