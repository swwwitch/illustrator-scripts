# Unified Type Panel

[![Direct](https://img.shields.io/badge/Direct%20Link-UnifiedTypePanel.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/UnifiedTypePanel.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/UnifiedTypePanel.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Setting up a single piece of text means bouncing between the Character panel, the Paragraph panel, Mojikumi Settings, and Kinsoku Settings. Change the font, fix the kerning, match the leading, and finally set the mojikumi — that round trip quietly eats time.

On top of that, the combinations you reach for ("body text is like this, headings are like that") are almost always the same. Rebuilding them by hand every time is not especially productive.

So this script collects the settings that matter for text composition into a single persistent palette, and **applies them to the selected text the moment you touch a control**.

<img alt="" src="" width="50%" />

## Usage

1. Select the target text (a text frame, text nested in a group, or a range selected in text-edit mode).
2. Run the script.

**There is no Apply button.** The setting is applied to the selected text as soon as you operate a radio button, slider, or input field. Press `Esc` to close the palette. Running the script again while the palette is already open simply brings it to the front instead of opening a second one.

Every time the palette regains focus it re-reads the current values of the selected text into the UI, so selecting different text updates the display along with it.

## Top area: Policy and Type

Two panels at the top of the palette apply frequently used combinations in a single click: Policy on the left, Type on the right.

### Policy

Three settings change depending on whether the text is Japanese, Roman, or mixed: character alignment, leading basis, and hyphenation. This panel switches all three together.

| Policy | Character alignment | Leading basis | Hyphenation |
| --- | --- | --- | --- |
| Japanese | Center | Embox top | OFF |
| Roman | Roman baseline | Roman baseline | ON |
| Mixed | Roman baseline | Embox top | ON |

### Type

| Type | Applied settings |
| --- | --- |
| Body | Metrics - Roman Only / leading 150% / justify (last line left) / solid mojikumi / weak kinsoku v2 |
| Heading | Metrics / leading 115% / left align / tight mojikumi / weak kinsoku v2 |

Clicking either panel switches automatically to the Composition tab, where the applied settings are reflected in the individual controls.

## Composition tab

This is the tab that opens at startup. Character attributes are on the left, paragraph attributes on the right.

### Auto kerning

| Option | Internal value |
| --- | --- |
| Metrics - Roman Only | METRICSROMANONLY (metrics for Roman text only) |
| 0 | NOAUTOKERN |
| Metrics | AUTO |
| Optical | OPTICAL |

Proportional metrics is turned ON automatically only when you choose "Metrics". Choosing anything else turns it back OFF. The checkbox at the top of the letter-spacing panel shows this linked state.

### Letter spacing

- **Tsume**: 0–100%
- **Tracking**: -100 to 500

Both can be operated from the input field or the slider. The slider steps by 1 normally, and **snaps to steps of 10 while you drag with Shift held down**. Values entered outside the range are clamped automatically.

In the input fields, `↑` and `↓` increment and decrement the value (Shift for steps of 10, Option for steps of 0.1).

### Character alignment

"Roman baseline" and "Center" are placed directly as radio buttons; the remaining four (embox top/right and bottom/left, ICF box top/right and bottom/left) are chosen from the popup to the right of the "Other" radio. Only the two you reach for most often are one click away.

### Font size and leading

This is the most distinctive part of the panel. **Leading is always set as auto leading.**

The panel has three input fields.

| Field | Content |
| --- | --- |
| Size | Font size |
| Leading | The effective leading value (size × leading %) |
| Leading % | The paragraph's auto-leading amount |

The percentage you enter is set as the paragraph's `autoLeadingAmount`, and auto leading is turned ON at the same time. In other words, Illustrator's Character panel always shows leading as "Auto", and **the leading follows automatically when you change the font size**. There is no need to reapply the leading.

If you type a pt value directly into the Leading field, the percentage is calculated back from the font size. The Leading field and the Leading % field both lead to the same result, whichever one you type into.

Note that right after the selection is re-read, the Leading field shows the **actual leading value of the selected text**, not the calculated size × %.

Units follow the document's unit setting (`text/units`). In a Q/H environment the Size field is labeled "Q" and the Leading field "H".

### Leading basis

A choice between embox top and Roman baseline. It switches only the basis, leaving the leading value itself unchanged.

### Justification

Left, center, right, justify (last line left), and justify all are selected with icon buttons. The icons are drawn by the script and switch between light and dark to match Illustrator's UI brightness setting (re-evaluated every time the palette regains focus).

Changing the justification can shift the visual position of the text frame. This panel compares `visibleBounds` before and after applying and moves the frame by the difference, so the justification changes **while the visual position stays put**.

### Japanese composition

Kinsoku and Mojikumi Settings are collected in a single panel, each chosen from a popup.

| Kinsoku | Mojikumi |
| --- | --- |
| None | None |
| Hard | Full/half-width punctuation at line end |
| Soft | Half-width punctuation |
| Soft v2 | Half-width punctuation at line end |
| | Full-width punctuation at line end |
| | Full-width punctuation |
| | Tight |
| | Solid |

**Choosing "None" for kinsoku has no effect.** Illustrator's scripting cannot set kinsoku to "None" (an action has to be played instead). The item is listed, but specifying it is ignored.

## Size tab

### Font size

Size, scale (horizontal and vertical set together as one value), and effective size (size × scale) are lined up. When the scale is 100% the effective size equals the actual size, so the "Effective" row is dimmed.

The Size field is synced with the Size field on the Composition tab; editing either one gives the same value.

### Apparent ←→ actual size

This button is a two-way conversion whose direction flips each time you press it. It rearranges the split between actual size and scale **while preserving the apparent size (size × scale)**.

- **When the scale is not 100%**: the current apparent size is baked into the actual size and the scale is set to 100% (apparent → actual).
- **When the scale is 100%**: pressing it after changing the size calculates the scale that restores the **apparent size from before the change** (actual → apparent).

For the latter, the panel keeps the "apparent size to preserve" internally. Editing the Size field does not update this reference, which is what makes it possible to return to the original appearance after changing the size. Editing the scale, or re-reading the selection, makes the apparent size at that moment the new reference.

### Type scale

Generates a list of seven sizes from a base size and a ratio. The ratio can be chosen from musical intervals.

| Ratio |
| --- |
| Minor Second 1.067 |
| Major Second 1.125 |
| Minor Third 1.2 |
| Major Third 1.25 (default) |
| Golden Ratio: ½ 1.309 |
| Perfect Fourth 1.333 |
| Augmented Fourth 1.414 |
| Golden Ratio 1.618 |

Step numbers in the list treat the base size as 0, with two steps above (-2, -1) and four below (1–4). Clicking a row applies that size to the selected text. When the base size field is empty, it is initialized with the current font size (or the equivalent of 12 pt if that cannot be read).

## Presets tab

Register a combination of font plus auto kerning, tsume, and tracking as a preset, and apply the whole thing with a single click.

| Button | Action |
| --- | --- |
| Add | Save the current selection's settings as a preset |
| Overwrite | Update the preset selected in the list with the current selection's settings |
| Delete | Remove the preset selected in the list |
| Export | Write the registered presets to a JSON file (Desktop by default) |

The list can be switched between "Font only" and "Details" (five columns: font, auto kerning, tsume, proportional, tracking). Presets are persisted as JSON under `Folder.userData`.

The list is first filled with PostScript names, which are replaced with the fonts' real names (the Japanese names for Japanese fonts) once the palette is shown.

## Document fonts tab

Scans the document for the combinations of "font plus settings" actually in use and lists them.

This list also has two display modes, "Font only" and "Details"; Details has eight columns: count, font, size, leading, auto kerning, tsume, tracking, and proportional.

**What a click does depends on the display mode.**

- **Font only**: clicking applies just that font to the selected text.
- **Details**: clicking applies the **entire combination** on that row (font, size, kerning, tsume, tracking, and so on) to the selected text.

The [Select text with the chosen font] button works in the opposite direction: it searches the document for text matching the conditions selected in the list and selects it. Here too, "Font only" matches on the font and "Details" matches on the whole combination.

The Filter panel narrows the scan.

| Item | Default |
| --- | --- |
| Include locked text | ON |
| Include hidden text | ON |
| Limit to the current artboard | OFF |

Scanning is expensive, so it runs only when you switch to this tab, when you change a filter, and when you press [Refresh list].

## Footer

| Button | Action |
| --- | --- |
| Hidden characters | Show or hide hidden characters (line breaks, spaces, and so on) |
| Reset | Restore 12 pt / scale 100% / tsume 0 / tracking 0 / Metrics / Roman baseline / left align / leading 115% / embox top / tight mojikumi / weak kinsoku v2, and set aki (before/after) back to auto |
| Reload | Re-read the current values of the selected text into the UI |

## The unit of application differs per item

This is easy to overlook: how far a setting reaches depends on the item.

| Item | Scope |
| --- | --- |
| Auto kerning / tsume / font size (via type scale) | The **entire paragraph** the selection touches |
| Tracking / font size (Size field) / scale / font | The **selected range** only |
| Justification / mojikumi / kinsoku / leading | Paragraph or frame |

For example, selecting a few characters in text-edit mode and changing the tsume applies to **the whole paragraph**, not just those characters, whereas tracking applies only to the selected range. Kerning and tsume tend to cause trouble when they vary within a paragraph, so they are deliberately kept at paragraph scope.

## Targets

Text frames, text nested in groups (processed recursively), and TextRange selections in text-edit mode.

## Notes

Illustrator's persistent palettes lose their connection to the document DOM while they are displayed, so all actual object manipulation is delegated to the main engine via BridgeTalk. The delegated code is wrapped with `encodeURIComponent` before being sent.

Left alignment is a special case. Illustrator sometimes ignores an assignment of `Justification.LEFT` from scripting, so it goes through a workaround that temporarily runs `resize(200,200)` → `resize(50,50)` on the frame to refresh the paragraph attributes (position and matrix are restored). It is a fallback for that known bug only, and is not used for right or center alignment.

Paragraphs whose kinsoku is "None" throw an exception (Error 9563) just from reading the value. Reading the current value is therefore wrapped in a try, and an exception is treated as "None".

The two lists on each tab (Font only / Details) are stacked at the same position with `orientation = "stack"`. A stack group's height is the maximum of its children, so placing two of them does not make the palette taller. Generating both before display also guarantees that `onChange` fires (in some environments, clicks do not register on a listbox created dynamically after display).

### note

- [An Illustrator palette script dedicated to text composition, unifying the Character, Paragraph, and OpenType panels (Japanese)](https://note.com/dtp_tranist/n/n4e2b79cf2891)
