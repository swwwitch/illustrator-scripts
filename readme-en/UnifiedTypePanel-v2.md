# UnifiedTypePanel-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-UnifiedTypePanel--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/UnifiedTypePanel-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/UnifiedTypePanel-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A docked palette that sets text-composition attributes (font, font size, auto kerning,
letter spacing, character alignment, justification, leading, and mojikumi) for the selected text.
A "Type" role box (Body/Heading and Japanese/Roman/Mixed) sits just below the info bar; below it
are four tabs ("Fonts": document fonts; "Presets": presets; "Type" = center: font size, kerning,
Tsume / right: character alignment, leading, basis; "Japanese Composition": justification,
kinsoku, mojikumi).

- Top info bar: when exactly one text is selected, shows font family, style, size, leading, justification, auto kerning, and kinsoku

- Document fonts: lists the fonts used in the document; clicking one applies it to the selection
- Presets: a font plus kerning / Tsume / tracking, applied together when clicked. "Add"
  saves the current selection's settings, "Overwrite" updates the selected preset, "Delete" removes it (persisted as JSON under Folder.userData)
- Font size: size, scale (horizontal/vertical set together), effective (size × scale, shown)
- Auto kerning: Metrics - Roman Only / 0 / Metrics / Optical (proportional metrics ON only for "Metrics")
- Letter spacing: Tsume (0–100%) and tracking (-100 to 500), via input fields and sliders (Shift = coarse steps)
- Character alignment: Roman baseline / center / Other (embox top-bottom & ICF box top-bottom via popup)
- Justification: left / center / right / justify (last left) / justify all (applied while keeping the text's visual position)
- Type: Body (solid mojikumi) / Heading (tight mojikumi); applies common combinations at once.
  Also Japanese (center align / top-to-top basis), Roman (Roman-baseline align / baseline basis),
  and Mixed (Roman-baseline align / top-to-top basis) as a one-shot alignment + leading-basis preset
- Leading: 110% / 125% / 150% / Other (enter a %) / Auto, individual vs common, and leading basis.
  Individual (per-line base size) vs common (the frame's most frequent size) is chosen with radios
  (Leading choices: 110% / 125% / 150% / Other / Auto)
  Choosing "Other" prefills the current leading as %; "Change auto value" edits the auto-leading percentage
- Mojikumi: None / half-width punctuation / full-width punctuation / tight / solid, applied together via radio buttons (on the "Japanese Composition" tab)
- Show/hide hidden characters, Reload (re-read the selection's current values), and
  Reset (restore defaults; kerning Metrics, mojikumi Tight, and aki before/after set to auto)
- Operating a radio or field applies it immediately to the current selection
- Selection handling covers not just standalone text frames but text nested in groups and ranges selected in text-edit mode (consistent across all features, including leading)
- Whenever the palette regains focus — or via "Reload" — the current selection's values are read back into the UI
- Runs as a persistent palette (#targetengine). The persistent engine's app
  loses its DOM connection while the palette is shown, so all DOM work is
  delegated to the main engine via BridgeTalk (code wrapped in encodeURIComponent)

### Script info

- Version: v1.1.0
