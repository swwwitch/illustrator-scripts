# ReplaceTextWithPasteSequential.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-ReplaceTextWithPasteSequential.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceTextWithPasteSequential.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceTextWithPasteSequential.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script takes only the first line of the multi-line text on the clipboard and applies it to the selected text frames.

The applied line is removed from the clipboard, so running the script repeatedly walks through the lines one by one. Blank lines are skipped, so a list with spaced-out entries works as it is. It is meant for filling in a name list or any other list, one entry at a time.

With no text frame selected, the first line is pasted at the center of the window and the remaining lines stay on the clipboard, so running the script repeatedly places the entries one by one. It is a handy way to put the first frame on the canvas.

It is derived from `ReplaceTextWithPaste.jsx`. That script applies the same text everywhere at once, while this one consumes a single line per run.

## Main features

- Applies the first line on the clipboard to the selected text frames
- Removes the applied line from the clipboard so the next run continues where the last one stopped
- Pastes the first line at the center of the window when nothing is selected, leaving the remaining lines on the clipboard
- Recognizes paragraph returns (CR), soft returns (LF), and CRLF as line breaks
- Skips blank lines, including whitespace-only ones, instead of emptying the text frame
- Walks into groups and clip groups, and processes the text frames inside them
- Reports when the last line has been applied
- Keeps the selection intact across the run
- Japanese and English UI

## Usage

1. Copy the multi-line text you want to distribute.
2. Select the text frame that should receive the first line.
3. Run `ReplaceTextWithPasteSequential.jsx`.
4. Select the next text frame and run it again. Repeat until no lines are left.

Assigning a keyboard shortcut lets you work through the list with nothing but select-and-run.

## Workflow

1. Save the current selection. When it holds no text frame at all, trim the pasted text frame down to its first line, move it to the center of the window, write the remaining lines back to the clipboard, and stop there; none of the steps below run.
2. Clear the selection, run a normal paste twice, then read the string from the text frame pasted the second time. (The first paste only refreshes Illustrator's cached clipboard; whatever it pastes is removed right away.)
3. Remove the pasted objects and restore the saved selection. When the paste arrives as a group, look for a text frame inside it.
4. Strip the blank lines from the string, then split it at the first line break into the first line and the remainder.
5. Replace the contents of the selected text frames with the first line.
6. Write the remainder back to the clipboard through a temporary text frame. Report and stop when there is no remainder.

## Scope

| | Objects |
| --- | --- |
| Handled | Text frames, and text frames inside groups and clip groups |
| Not handled | Images, shapes, locked objects |

## Notes

- When nothing is selected — including a selection that holds no text frame — the pasted text frame is trimmed down to its first line, placed at the center of the window, and the remaining lines are written back to the clipboard. A run consumes exactly one line, just as it does when text frames are selected.
- When the paste holds no text frame at all (an image or a shape, for instance), it is placed at the center of the window as it is, with no line taken and no clipboard write-back.
- Placement is based on the center of the current window (`activeView.centerPoint`). Illustrator pastes there too, but trimming down to the first line changes the size of the frame, so it is measured again and re-centered afterwards. When several objects arrive at once, they are all moved by the same delta so their layout is preserved.
- When several text frames are selected, all of them receive the same first line. A run always consumes exactly one line.
- The clipboard is left untouched after the last line is applied. Running the script again in that state reapplies the same line.
- The write-back uses Illustrator's own copy, so the clipboard ends up holding an Illustrator object. Pasting it into another application in the middle of a run gives you artwork rather than text.
- If the copied source is a set of separate single-line text frames rather than one multi-line frame, only the first one is read.
- Blank lines, including whitespace-only ones, are skipped. They are also dropped from the clipboard that gets written back, so later runs never meet them either.
- Illustrator holds on to whatever it copied itself, so the first paste after another application changes the clipboard still brings back the old contents. To work around this, the script pastes twice and uses the result of the second paste.
- Alignment and character styles of the original text frame are preserved; only the contents are replaced.
- Illustrator has no undo-grouping API, and the script also creates and removes a temporary text frame, so undoing a run takes multiple steps.
- Original idea by Gorolib Design

## Changelog

- v1.0.3 (20260825): The paste with no selection now places only the first line instead of the whole clipboard, and writes the remaining lines back to the clipboard, consuming one line per run just as it does with a selection. Placement is now based on the center of the window (`activeView.centerPoint`) rather than the center of the artboard
- v1.0.2 (20260825): Added the behavior for running with no text frame selected. Instead of reporting and stopping, the script now pastes the clipboard at the center of the artboard as it is, without consuming a line or writing the clipboard back
- v1.0.1 (20260816): Fixed text inside a selected group or clip group sometimes not being replaced. The walk into groups now runs before the paste, so the targets are collected while the references are still valid. When the selection holds no text frame at all, the script now reports it and stops instead of consuming a line from the clipboard
- v1.0.0 (20260814): Initial release
