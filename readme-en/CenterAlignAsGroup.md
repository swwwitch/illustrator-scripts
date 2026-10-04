# Center the selection while keeping its layout


[![Direct](https://img.shields.io/badge/Direct%20Link-CenterAlignAsGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/single-function/CenterAlignAsGroup.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterAlignAsGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Temporarily groups the selected objects, then runs Align Horizontal Centers and Align Vertical Centers from the Align panel. The selection keeps its internal spacing and moves to the center as one piece. When the selection is a single text object holding one line, its justification is changed to centered as well.

Align panel commands cannot be called from the DOM, so the script writes an action definition to a temporary file, loads it, runs it, and discards it right away — the "dynamic action" approach.

### Features

- Groups the objects temporarily when two or more are selected, and ungroups them afterwards
- Switches to the artboard holding the selection when it sits outside the current one
- Sets paragraph justification to centered when a single one-line text object is selected
- Pulls each line of centered point type toward its glyph center with kerning at the line start
- Turns on Align to Glyph Bounds (point type and area type) for the run only, then restores the previous state
- Promotes a text selection made with the Type tool to the text object itself
- Stops with an alert when the selection spans multiple layers
- Always removes the loaded action, so nothing is left behind in the Actions panel

### How it works

1. Check that a document is open and something is selected
2. If characters are selected, reselect the text object instead
3. Check that the selection does not span multiple layers
4. Make the artboard holding the selection the active one
5. Save the current Align to Glyph Bounds state and turn it on
6. Center the justification when the selection is a single one-line text object
7. For centered point type, kern the start of each line to pull its glyphs toward the center
8. Group (when more than one object is selected) → center with the action → ungroup
9. Restore Align to Glyph Bounds

### Notes

- The alignment reference (selection / key object / artboard) follows the Align panel setting. **With "Align to Selection", grouping leaves a single object, so nothing moves.** Set the reference to the artboard or a key object before running the script.
- Multi-layer selections are rejected because grouping moves every object to the layer of the frontmost one, and ungrouping does not send them back.
- Grouping and ungrouping add two extra undo steps.
- Running it on a character selection inside threaded text targets every text object of that story.
- Justification is only changed for one-line text. Recentering a multi-line block would move every line and change its look; area type that wraps onto two or more lines is left alone for the same reason.
- Centered point type is measured line by line, and kerning at the line start pulls the glyphs toward the center (35% strength). This keeps lines from looking off because of the space in a trailing 、 or a leading 「. Area type and type on a path are left alone.
- In a document with several artboards, the one overlapping the selection most is used. If the selection overlaps none of them, the artboard nearest its center is used instead.
- To center vertically only, use [VerticalCenterAlignAsGroup](VerticalCenterAlignAsGroup.md).

### Update history

- v1.1.0 (20261003) : Centered point type now kerns the start of each line so lines that look off because of trailing punctuation such as 、 are pulled toward the glyph center
- v1.0.3 (20260928) : Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure
- v1.0.3 (20260928) : Clip groups are measured by their masks when finding the artboard the selection is on
- v1.0.1 (20260821) : Added the automatic switch to the artboard holding the selection, and centered justification for one-line text
- v1.0.0 (20260821) : Initial release
- v1.1.1 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
