# KPTSketchy

[![Direct](https://img.shields.io/badge/Direct%20Link-KPTSketchy.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/KPTSketchy.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KPTSketchy.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Applies randomized transformations to the selected objects for a hand-drawn / sketchy look.
- Supports fill/stroke splitting, round corners, offset path, roughen (jagged / distortion) and grouping.
- Everything is applied as a live effect, so it stays editable and removable afterwards.
- Runs as a persistent palette, so the document stays usable while it is open.

### Usage

1. Run the script to show the palette.
2. Select objects and adjust the panels (the result is rebuilt on every change).
3. Press Recalculate to reroll the random values; just leave it as is to keep it.
4. Leaving the palette, or closing it (X / Esc), finalizes the current result.

### Notes

- Round corners, offset and move use the current ruler unit (rulerType).
- Effect order: Round Corners -> Offset Path -> Transform -> Roughen (Distortion -> Jagged).
- Offset Path is applied twice, negative first and then positive.
- Only objects having both a visible fill and a visible stroke can be split.
- Groups are expanded into their children when "Process group items individually" is on.
- Clipping paths are excluded from processing.
- Changing a value or recalculating undoes the previous result before applying a new one.
- Once the palette loses focus the result is finalized and never undone again,
  so that operations the user performed in the document are never reverted by mistake.

### Implementation notes

- A palette's app object loses its document connection, so every DOM operation lives in a
  worker function and is delegated to the main engine through BridgeTalk.
- Worker functions are concatenated with toString(), encoded with encodeURIComponent and
  sent as eval(decodeURIComponent(...)).
- Worker functions must avoid line comments and must terminate every statement with a semicolon.

### Script info

- Version: v1.2.0
- First release: 2026-04-14
- Last updated: 2026-07-22
