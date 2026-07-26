# ReplaceTextWithPaste.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-ReplaceTextWithPaste.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceTextWithPaste.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceTextWithPaste.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script replaces the contents of the selected text frames with the text on the clipboard.

With nothing selected, it creates a new text frame where the text was copied from. When a group is selected, every text frame inside it is processed.

## Main features

- Replaces the contents of every selected text frame at once
- Creates a new text frame at the original position when nothing is selected
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
2. Run a normal paste once, then read the contents and bounds from the pasted text frame.
3. Remove the pasted objects and restore the saved selection.
4. Replace the contents of the selected text frames, or create a new one when nothing was selected.

## Scope

| | Objects |
| --- | --- |
| Handled | Text frames, and text frames inside groups |
| Not handled | Images, shapes, locked objects |

## Notes

- The script performs a normal paste internally. Nothing happens when the clipboard holds no text.
- Alignment and character styles of the original text frame are preserved; only the contents are replaced.
- Illustrator has no undo-grouping API, so undoing this run takes multiple steps.
- Original idea by Gorolib Design

## Changelog

- v1.0.0 (20260727): Introduced the version variable, removed the ineffective `doc.undoGroup` assignment, consolidated the duplicated selection-restore and clipboard-read logic, and added a message for when no document is open
