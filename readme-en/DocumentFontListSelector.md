# DocumentFontListSelector

[![Direct](https://img.shields.io/badge/Direct%20Link-DocumentFontListSelector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/DocumentFontListSelector.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DocumentFontListSelector.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A docked palette that lists the text-composition combinations used in the
document (font, size, leading, Tsume, tracking, auto-kerning and proportional
metrics), deduped. Even for the same font, a different combination of these is
listed as a separate candidate.

The list is shown as eight columns: Count / Font / Size / Leading / Auto Kerning /
Tsume / Tracking / Prop. Metrics (proportional metrics shown as ON/OFF). Count is
the number of text frames that use the combination (a frame counts once).

- The palette is a centered row of top checkboxes ("Include locked text" /
  "Include hidden text" / "Current artboard only"), the column list, and
  a button area below it (left: an "Apply on click" checkbox / center: a spacer /
  right: "Refresh list" and "Select matching text" buttons)
- While "Apply on click" is ON, clicking a row applies that combination (font,
  size, leading, auto kerning, Tsume, tracking, proportional metrics) to the
  current selection; while OFF, clicking does not apply
- The "Select matching text" button selects the text frames containing
  characters that match the selected condition (this turns "Apply on click" OFF)
- The "Refresh list" button rescans the list on demand
- Turning on "Current artboard only" limits the scan and matching selection to
  text frames overlapping the active artboard (off = whole doc); the apply target
  is always the current selection
- While "Include hidden text" is OFF, hidden text (hidden items/groups, or items
  on hidden layers) is excluded from the scan and matching selection (ON includes it)
- While "Include locked text" is OFF, locked text (locked items/groups, or items
  on locked layers) is excluded from the scan and matching selection (ON includes it)
- The list is sorted by font, then size, leading, auto kerning, Tsume, tracking,
  and prop. metrics (not by count)
- The list is rescanned when shown, on "Refresh list", and when a checkbox is toggled
- Runs as a persistent palette (#targetengine). The persistent engine's app
  loses its DOM connection while the palette is shown, so all DOM work is
  delegated to the main engine via BridgeTalk (code wrapped in encodeURIComponent)

### Script info

- Version: v1.1.3
