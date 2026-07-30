# AiTextOutlineRestorePalette.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-AiTextOutlineRestorePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/outline/AiTextOutlineRestorePalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiTextOutlineRestorePalette.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

Sometimes, after text has been outlined for delivery, someone asks to fix just one character.

Once text is outlined, all that is left is a set of paths. Which font was it? What size, what leading? What about kerning? Reopening the original file would answer that, but the original file is not always at hand.

This script is a persistent palette that **writes the attributes into the object's note right before outlining, and reads that note later to turn the paths back into text**. As long as you have the outlined file, you can restore from it.

## Usage

### Outlining

1. Select the text.
2. Click **Outline with Note** in the palette.

The result is the same as a normal outline, except that the attributes are stored in the resulting object's note. What was saved is listed in the palette right away.

### Restoring

1. Select an outline (path or group) that carries a note.
2. Click **Restore Text**.

The note is parsed and the text frame is rebuilt.

The palette can also be closed with the `Esc` key.

## Restored properties

| Item | Notes |
| --- | --- |
| Text | Multi-line supported |
| Orientation | Vertical / horizontal |
| Font | Saved as the PostScript name |
| Font size | |
| Leading | |
| Auto leading | When on, the leading value is not written and leading returns to automatic |
| Horizontal / vertical scale | |
| Kerning | Metrics / Optical / Metrics (Roman Only) / None |
| Proportional metrics | Follows the kerning method |
| Tracking | |
| Tsume | |
| Alignment | Left / center / right / justify (including how the last line is handled) |
| Kinsoku | Saved as the set name |
| Mojikumi | Saved as the set name |
| Fill color | CMYK / RGB / Gray / Spot |
| Coordinates | Based on geometricBounds; not shown in the list |

## What cannot be restored

This is the part that is easy to misread, so here is a summary.

**Only the first character's values are saved.** If color or size varies inside one text frame, the whole frame is restored with the first character's values. Per-character differences do not come back.

**Gradient and pattern fills are not supported.** They are saved as empty and the restored text keeps the default color (this is deliberate, so the text never becomes invisible).

**A kinsoku setting of "none" is skipped on restore.** Illustrator does not allow assigning "none" to kinsoku from a script. Paragraphs that had no kinsoku keep the default kinsoku after restore. Mojikumi, on the other hand, can be restored as "none".

**Missing fonts fall back to the default font.** The status line then shows "Some fonts used defaults". The other attributes are still restored; the process does not stop.

**Japanese-only attributes are not handled in English locales.** Orientation, kinsoku, mojikumi and tsume are neither saved, displayed nor restored.

**Older notes skip only the items they do not record.** Outlines created by earlier versions can still be restored, as far as their notes go.

## Restore options

| Option | On (default) | Off |
| --- | --- | --- |
| Keep outline data | The original outlines are moved to the `outlined_text` layer | The outlines are deleted after restore and no `outlined_text` layer is created |
| Restore text to a separate layer | Text is restored to the `restored_text` layer | Text is restored to the same layer as the outline and no `restored_text` layer is created |

The stashed outlines are set to 30% opacity and locked. The layer itself is sent to the back, given the template-layer attribute, and locked, so it does not get in the way of the restored text.

The `outlined_text` layer is not created again and again: an existing one is unlocked and reused, so all outlines are collected on a single layer.

## The two buttons

- **Load Note**: loads the selected object's note and shows it in the list
- **Attributes**: toggles Illustrator's Attributes panel (for viewing or editing the note directly)

The selected object's note is also loaded automatically when the palette opens. With several objects selected, the **first object that carries a note** is shown.

## Targets

- Outlining: TextFrame
- Restoring: PathItem, GroupItem (those carrying a note)

## Notes

Notes are saved with Japanese labels for compatibility. In English locales, item names and enumerated values are converted to English just before display. Numbers, font names and color values pass through unchanged.

The internal name of a mojikumi set varies between environments (spacing, separators, case), so names are normalized before being compared. When that still fails, a fallback resolves the internal romanized name by prefix match.

Because a persistent palette in Illustrator loses its connection to the DOM while it is displayed, all object manipulation is delegated to the main engine via BridgeTalk.

The template-layer attribute cannot be set through the API, so a temporary action is loaded and played instead. The action is unloaded when the palette is closed.

## Article

[Restoring outlined text in Illustrator (Japanese)](https://note.com/dtp_tranist/n/nc476be8ad43c)

## Update history

- v2.0.1 (2026-07-31): Moved the overview into the README, tidied the basic info block, updated the article URL
- v1.0 (2024-07-23): Initial version
- v1.9 (2026-07-04): Persistent palette, text restore, note display, kerning and fill color restore, multi-line support, robustness
- v1.10 (2026-07-05): Added saving and restoring of alignment, the auto-leading flag, horizontal/vertical scale, kinsoku and mojikumi
- v2.0.0 (2026-07-05): Japanese labels for kinsoku and mojikumi in the listbox, reordered display, `outlined_text` layer reuse, new title, localized listbox item names and values, split Outline and Restore Text panels, added the Keep outline data and Restore text to a separate layer options, helpTips on panels and listbox, unified note wording, and Japanese-only attributes skipped in English locales
