# KPTSketchy

[![Direct](https://img.shields.io/badge/Direct%20Link-KPTSketchy.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/KPTSketchy.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KPTSketchy.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Applies randomized transformations to the selected objects for a hand-drawn / sketchy look.
- Supports fill/stroke splitting, round corners, offset path, Roughen (fine / coarse) and grouping.
- Everything is applied as a live effect, so it stays editable and removable afterwards.
- Runs as a persistent palette, so the document stays usable while it is open.

### Usage

1. Run the script to show the palette.
2. Select objects and adjust the panels (the result is rebuilt on every change).
3. Press Reroll to draw new random values; just leave it as is to keep it.
4. Leaving the palette, or closing it (X / Esc), finalizes the current result.

### Panels

| Panel | Contents |
| --- | --- |
| Apply to | Whether to process group items individually, split fill and stroke, and group the split pair |
| Random transform | Ranges for scale, move and rotate |
| Corners & Offset | Corner radius and offset amount |
| Roughen: Fine | Adds fine jitter to the outline for a pen-drawn look |
| Roughen: Coarse | Warps the outline on a larger scale to distort the shape itself |

### Notes

- Round corners, offset and move use the current ruler unit (rulerType).
- The three Random transform fields are plus-minus ranges; the `±` in the unit label marks them as such.
- Roughen: Fine and Roughen: Coarse are both the Roughen effect. Its Points option is pinned to
  Smooth, so the two differ only in size and detail.
- Effect order: Round Corners -> Offset Path -> Transform -> Roughen (Coarse -> Fine).
- Offset Path is applied twice, negative first and then positive.
- Only objects having both a visible fill and a visible stroke can be split.
- Groups are expanded into their children when "Process group items individually" is on.
- Clipping paths are excluded from processing.
- Negative entries are clamped to 0 and detail fields are rounded to an integer of at least 1;
  the field itself is rewritten so it shows the value that is actually applied.
- Input that is not a number (`3abc` and the like) is not applied, and the status line says so.
- Changing a value or recalculating undoes the previous result before applying a new one.
- Once the palette loses focus the result is finalized and never undone again,
  so that operations the user performed in the document are never reverted by mistake.
- After a timeout it is unknown whether the effects were applied, so the automatic undo stops.
  Undo manually if the effects ended up applied twice.

### Implementation notes

- A palette's app object loses its document connection, so every DOM operation lives in a
  worker function and is delegated to the main engine through BridgeTalk.
- Worker functions are concatenated with toString(), encoded with encodeURIComponent and
  sent as eval(decodeURIComponent(...)).
- Worker functions must avoid line comments and must terminate every statement with a semicolon.
- The palette reference is kept on `$.global` so the persistent engine holds it between runs.
  A function-scoped `var` is hoisted as `undefined`, so the common
  `var x = (typeof x !== "undefined") ? x : null` idiom can never carry a value over.

### Article

[Give artwork a hand-drawn, slightly rough look with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/na808bac430d9)

### Changelog

- v1.2.1 (2026-09-05): Keep the palette reference on `$.global` so a re-run reliably closes the
  previous palette, revise the automatic undo after a timeout or a failed send, validate and
  normalize numeric input, revise the UI wording (Roughen panel names, Reroll and others),
  and clean up naming and code structure
- v1.2.0 (2026-07-22): Turned into a persistent palette, DOM work delegated to the main engine
  through BridgeTalk, rebuilt as a two-column UI
- v1.1.0 (2026-04-14): Added Transform (scale / move / rotate) and Offset Path
- v1.0.0 (2026-04-14): First release

---

### Script info

- Version: v1.2.1
- First release: 2026-04-14
- Last updated: 2026-09-05
