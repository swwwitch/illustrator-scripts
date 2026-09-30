# Replace the selection with the clipboard contents

[![Direct](https://img.shields.io/badge/Direct%20Link-ReplaceWithPaste.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceWithPaste.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceWithPaste.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script replaces the contents of the selected text frames with the text on the clipboard.

With nothing selected, it creates a new text frame with the default formatting where the paste lands, which is the center of the view. When a group is selected, every text frame inside it is processed.

When the clipboard holds something other than text (shapes, images, and so on), each selected object is replaced with it. The pasted contents are centered on the original object, and you choose the sizing: Keep Size, Fit Long Side, or Fit Short Side (fitting scales proportionally).

To fill in the lines one at a time instead, use the derived [ReplaceTextWithPasteSequential.jsx](ReplaceTextWithPasteSequential.md). That one pastes the clipboard at the center of the artboard when nothing is selected.

## Main features

- Replaces the contents of every selected text frame at once
- Creates a new text frame with the default formatting where the paste lands — the center of the view — when nothing is selected
- Walks into groups and clip groups, and processes the text frames inside them
- When the clipboard holds something other than text, replaces the selected objects with it (keep size, fit the long side, or fit the short side; centered and kept in the same stacking position)
- Supports point type, area type, and type on a path
- Keeps the selection intact across the run
- Japanese and English UI

## Usage

1. Copy the replacement text or objects.
2. Select the text frames or objects to replace (for text, leave nothing selected to create a new frame).
3. Run `ReplaceWithPaste.jsx`.

## Workflow

1. Save the current selection.
2. Clear the selection, run a normal paste twice, then read the contents and bounds from the text frame pasted the second time. (The first paste only refreshes Illustrator's cached clipboard; whatever it pastes is removed right away.)
3. Remove the pasted objects and restore the saved selection. When the paste arrives as a group, look for a text frame inside it.
4. When text is found, replace the contents of the selected text frames, or create a new one when nothing was selected.
5. When no text is found, paste again for each selected object, size the paste as chosen, and remove the original. With nothing selected, paste at the center of the view as usual.

## Scope

| | Objects |
| --- | --- |
| Handled (text copied) | Text frames, and text frames inside groups and clip groups |
| Handled (non-text copied) | The selected objects, text frames included |
| Not handled | Locked objects |

## Settings

Switch these in the User Settings block at the top of the script. Both apply only when replacing with non-text contents.

| Variable | Values | Meaning |
| --- | --- | --- |
| `SHOW_SIZE_DIALOG` | `true` (default) / `false` | Whether to open a dialog to choose the sizing. Its Show Preview checkbox (on by default) shows the result of the chosen sizing on the spot |
| `DEFAULT_SIZE_MODE` | `"keep"` / `"long"` (default) / `"short"` | Keep size / fit the long side / fit the short side. Used as is when the dialog is off, and as the initial choice when it is on |

To run from a keyboard shortcut without the dialog, set `SHOW_SIZE_DIALOG` to `false`.

## Notes

- The script performs a normal paste internally. It reports and stops when nothing gets pasted.
- While editing selected characters, it reports and stops without replacing anything if the clipboard holds something other than text.
- With only a caret placed in text (no characters selected), the clipboard is inserted at the caret with Paste without Formatting (nothing is replaced).
- When replacing with non-text contents, a selected group is replaced as a whole (only text copies reach the text frames inside groups).
- Illustrator holds on to whatever it copied itself, so the first paste after another application changes the clipboard still brings back the old contents. To work around this, the script pastes twice and uses the result of the second paste.
- A new text frame lands wherever Illustrator pastes (the center of the view), not at the coordinates it was copied from.
- A new text frame uses the default formatting. The font and size of the copied text are not carried over.
- Alignment and character styles of the original text frame are preserved; only the contents are replaced.
- Illustrator has no undo-grouping API, so undoing this run takes multiple steps.
- Original idea by Gorolib Design

## Changelog

- v2.0.7 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v2.0.6 (20260930): Button rows with only right-side buttons are now centered
- v2.0.5 (20260930): Fixed an error when running with characters selected by the Type tool
- v2.0.4 (20260929): Dialog opacity changed to 98%
- v2.0.3 (20260928): The button row is now built with the shared part. Clip groups are now measured by their mask
- v2.0.2 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v2.0.1 (20260927): With only a caret placed in text (no characters selected), the clipboard is now inserted at the caret with Paste without Formatting instead of replacing the whole text frame
- v2.0.0 (20260927): Renamed from `ReplaceTextWithPaste.jsx` to `ReplaceWithPaste.jsx`. Added support for non-text clipboard contents: each selected object is replaced with the pasted contents, matched in center and stacking position (with nothing selected, the contents are pasted as usual). Sizing is Keep Size, Fit Long Side, or Fit Short Side, chosen in a dialog with a live preview or fixed with the `SHOW_SIZE_DIALOG` and `DEFAULT_SIZE_MODE` user settings
- v1.1.4 (20260825): Spelled out in the overview that a new text frame lands at the center of the view, and noted that the derived `ReplaceTextWithPasteSequential.jsx` now pastes the clipboard at the center of the artboard when nothing is selected (the behavior of this script is unchanged)
- v1.1.3 (20260816): Fixed text inside a selected group or clip group sometimes not being replaced. The walk into groups now runs before the paste, so the targets are collected while the references are still valid, and the frames are all collected before any of them is rewritten
- v1.1.1 (20260814): Fixed text copied in an application other than Illustrator not coming through. Illustrator holds on to whatever it copied itself, so the first paste after another application changes the clipboard brings back the old contents; the script now discards that first paste and uses the result of a second one. It also looks for a text frame inside a pasted group, and reports when nothing was pasted at all
- v1.1.0 (20260814): Clear the selection before pasting, fixing a case where the original selection — or the characters selected with the Type tool — could be deleted when the paste did not go through. Added a message for when no text is found on the clipboard. Replacement errors are now deduplicated and reported in a single alert instead of one per selected object. Corrected the description of where a new text frame is placed
- v1.0.0 (20260727): Introduced the version variable, removed the ineffective `doc.undoGroup` assignment, consolidated the duplicated selection-restore and clipboard-read logic, and added a message for when no document is open
