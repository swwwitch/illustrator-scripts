# Save type settings as presets and apply them to the selected text in one click

[![Direct](https://img.shields.io/badge/Direct%20Link-FontPresetPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FontPresetPicker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontPresetPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that keeps the fonts you reach for most, together with their size, leading, auto kerning, Tsume, tracking, character alignment, justification, kinsoku and mojikumi, and applies the whole set to the selected text with a single click in the list.

You can narrow down which groups of settings get applied, and reset horizontal/vertical scale, aki before and after, and baseline shift to their default state at the same time.

It is the "Favorites" tab of UnifiedTypePanel.jsx, carved out into a standalone script.

### Features

- Lists your saved presets; clicking one applies the whole set to the selected text at once
- A font list on top, and below it the selected preset's settings in two columns (label / value) — font size, leading, leading (%), leading basis, auto kerning, Tsume, proportional metrics, tracking, character alignment, justification, kinsoku and mojikumi
- A **Settings to Apply** panel picks which groups get applied: font, font size, leading, kerning/Tsume/tracking, alignment/kinsoku/mojikumi
- A **Clear** panel resets horizontal/vertical scale, aki before and after, and baseline shift as the preset is applied
- In either panel, Option (Alt)-click switches between "just this one" and "all of them"
- Groups left out are dimmed in the lower detail list
- Leading is saved as an auto-leading amount (%) rather than points, and always applied as auto leading, so it follows the font size
- Stays open while you work; Add and Overwrite switch on and off as you change the selection
- **Add** saves the selected text's current settings as a preset (an existing entry for the same font is overwritten)
- **Overwrite** updates the preset selected in the list with the selected text's current settings
- **Delete** removes the preset selected in the list
- Fonts are listed by display name (family + style), not by PostScript name

### Usage

1. Run the script to open the palette.
2. Select the text you want to restyle (a text object, or a range of characters).
3. Click a preset in the top list — it is applied to the selection right away.
4. The lower list shows what that preset holds.
5. Close the palette when you are done (Esc closes it too).

To save the current settings, select that text and press **Add**.
The palette stays open, so you can keep changing the selection and applying presets.

### Options

- **Settings to Apply** — only the ticked groups are applied; whatever is unticked is left as it already is in the text.
  - **Font** — the typeface
  - **Font Size** — the font size
  - **Leading** — the auto-leading amount (%) and the leading basis
  - **Kerning, Tsume & Tracking** — auto kerning, Tsume and tracking
  - **Alignment, Kinsoku & Mojikumi** — character alignment, justification, kinsoku and mojikumi
  - Option (Alt)-click a box to switch between just that one and all of them.
- **Clear** — the ticked items are reset to their default state as the preset is applied. They are never stored in a preset.
  - **Horizontal & Vertical Scale** — back to 100%
  - **Aki Before & After** — back to Auto (which is not the same as 0, meaning no aki)
  - **Baseline Shift** — back to 0
  - Option (Alt)-click works here too.
- **The lower detail list** — read-only; it shows what the preset selected above holds, in two columns. Clicking it does nothing.
- **Proportional Metrics** — ON only for presets whose auto kerning is Metrics. It follows the kerning method and cannot be set on its own.
- **Leading / Leading (%)** — the preset stores the percentage (the auto-leading amount); the Leading row shows the effective value, font size × percentage, for reference. Being relative to the size, the leading follows whenever the size changes.

### Notes

- With nothing selected you can still use **Delete**, but **Add** and **Overwrite** are disabled and clicking the list applies nothing.
- The selection is re-read when the palette opens and whenever it comes back to the front, so click the palette once after changing the selection.
- Applying and reading are delegated to Illustrator's main engine over BridgeTalk, so there is a slight lag between the click and the result.
- Because the palette is persistent, close it and run the script again after editing the code — otherwise the old code keeps running.
- The palette's position is saved to `FontPresetPicker_location.txt` each time it moves, and restored on the next launch.
- Presets are stored in `FontPresetPicker_presets.json`, directly under `Folder.userData`. This is a separate file from UnifiedTypePanel.jsx's `UnifiedTypePanel_presets.json`.
- A corrupt preset file is moved aside to `.bak` before the defaults are restored.
- If a preset's font is not installed, the list shows its PostScript name and clicking it applies no font (the other settings still apply).
- Justification, kinsoku and mojikumi are paragraph settings. Applying a preset to part of a line still changes the whole paragraph.
- Kinsoku "None" cannot be set from a script (an Illustrator limitation), so a preset set to None leaves the text's kinsoku untouched.
- Mojikumi is looked up in the document's `mojikumiSet`. In a document with custom sets, a preset may land on a different set than expected.
- The "line-end punct half" mojikumi set reads back as the same value as "half-width punctuation", so the two cannot be told apart. **Add** on such text records it as "half-width punctuation".
- **Clear**'s "Aki Before & After" restores Auto (`-1` in the DOM). It never sets `0`, which is the distinct state meaning no aki.
- Changing the justification moves point text on the canvas, because its anchor changes.
- Size and leading % are stored to two decimal places, so fractional values such as 10.5 pt are kept as they are.
- Leading is never baked in as points; it goes in as the auto-leading amount (`autoLeadingAmount`), so Illustrator's Character panel always reads "Auto".
- The auto-leading amount is a paragraph setting. Applying a preset to part of a line still changes the leading of the whole paragraph.
- When you **Add** from text whose leading is a fixed point value, the percentage is back-calculated from the font size.

### Article

https://note.com/dtp_tranist/n/n3d7f8b58ef88

### Release Notes

- v1.0.0 (20260917) : Initial release. The "Favorites" tab of UnifiedTypePanel.jsx, carved out as a persistent palette

### Script Info

- Version: v1.0.0
