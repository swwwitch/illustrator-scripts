# ReplaceTextWithPaste.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-ReplaceTextWithPaste.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceTextWithPaste.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceTextWithPaste.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script replaces the contents of the selected text frames with the text on the clipboard.

With nothing selected, it creates a new text frame with the default formatting where the paste lands. When a group is selected, every text frame inside it is processed.

To fill in the lines one at a time instead, use the derived [ReplaceTextWithPasteSequential.jsx](ReplaceTextWithPasteSequential.md).

## Main features

- Replaces the contents of every selected text frame at once
- Creates a new text frame with the default formatting where the paste lands when nothing is selected
- Walks into groups and processes the text frames inside them
- Supports point type, area type, and type on a path
- Keeps the selection intact across the run
- Japanese and English UI

## Usage

1. Copy the replacement text in a text editor or elsewhere.
2. Select the text frames to replace (leave nothing selected to create a new one).
3. Run `ReplaceTextWithPaste.jsx`.

## Workflow

1. Save the current selection.
2. Clear the selection, run a normal paste twice, then read the contents and bounds from the text frame pasted the second time. (The first paste only refreshes Illustrator's cached clipboard; whatever it pastes is removed right away.)
3. Remove the pasted objects and restore the saved selection. When the paste arrives as a group, look for a text frame inside it.
4. Replace the contents of the selected text frames, or create a new one when nothing was selected.

## Scope

| | Objects |
| --- | --- |
| Handled | Text frames, and text frames inside groups |
| Not handled | Images, shapes, locked objects |

## Notes

- The script performs a normal paste internally. It reports and stops when no text is found on the clipboard.
- Illustrator holds on to whatever it copied itself, so the first paste after another application changes the clipboard still brings back the old contents. To work around this, the script pastes twice and uses the result of the second paste.
- A new text frame lands wherever Illustrator pastes (the center of the view), not at the coordinates it was copied from.
- A new text frame uses the default formatting. The font and size of the copied text are not carried over.
- Alignment and character styles of the original text frame are preserved; only the contents are replaced.
- Illustrator has no undo-grouping API, so undoing this run takes multiple steps.
- Original idea by Gorolib Design

## Changelog

- v1.1.1 (20260814): Fixed text copied in an application other than Illustrator not coming through. Illustrator holds on to whatever it copied itself, so the first paste after another application changes the clipboard brings back the old contents; the script now discards that first paste and uses the result of a second one. It also looks for a text frame inside a pasted group, and reports when nothing was pasted at all
- v1.1.0 (20260814): Clear the selection before pasting, fixing a case where the original selection — or the characters selected with the Type tool — could be deleted when the paste did not go through. Added a message for when no text is found on the clipboard. Replacement errors are now deduplicated and reported in a single alert instead of one per selected object. Corrected the description of where a new text frame is placed
- v1.0.0 (20260727): Introduced the version variable, removed the ineffective `doc.undoGroup` assignment, consolidated the duplicated selection-restore and clipboard-read logic, and added a message for when no document is open
