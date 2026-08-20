# CenterAlignAsGroup


[![Direct](https://img.shields.io/badge/Direct%20Link-CenterAlignAsGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/CenterAlignAsGroup.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterAlignAsGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Temporarily groups the selected objects, then runs Align Horizontal Centers and Align Vertical Centers from the Align panel. The selection keeps its internal spacing and moves to the center as one piece.

Align panel commands cannot be called from the DOM, so the script writes an action definition to a temporary file, loads it, runs it, and discards it right away — the "dynamic action" approach.

### Features

- Groups the objects temporarily when two or more are selected, and ungroups them afterwards
- Turns on Align to Glyph Bounds (point type and area type) for the run only, then restores the previous state
- Promotes a text selection made with the Type tool to the text object itself
- Stops with an alert when the selection spans multiple layers
- Always removes the loaded action, so nothing is left behind in the Actions panel

### How it works

1. Check that a document is open and something is selected
2. If characters are selected, reselect the text object instead
3. Check that the selection does not span multiple layers
4. Save the current Align to Glyph Bounds state and turn it on
5. Group (when more than one object is selected) → center with the action → ungroup
6. Restore Align to Glyph Bounds

### Notes

- The alignment reference (selection / key object / artboard) follows the Align panel setting. **With "Align to Selection", grouping leaves a single object, so nothing moves.** Set the reference to the artboard or a key object before running the script.
- Multi-layer selections are rejected because grouping moves every object to the layer of the frontmost one, and ungrouping does not send them back.
- Grouping and ungrouping add two extra undo steps.
- Running it on a character selection inside threaded text targets every text object of that story.
- To center vertically only, use [VerticalCenterAlignAsGroup](VerticalCenterAlignAsGroup.md).

### Update history

- v1.0.0 (20260821) : Initial release
