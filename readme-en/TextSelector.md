# Select text frames by a combination of conditions

[![Direct](https://img.shields.io/badge/Direct%20Link-TextSelector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextSelector.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextSelector.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Selects text frames across the document by a combination of conditions.

### Features

- Font Attributes: checkboxes for family, style, font size and fill color, using the selected text as the reference (every checked attribute must match)
- Text Type: point text, area text and text on a path (checkboxes, all on by default)
- String: none, exact, partial, prefix, suffix or regular expression (can be combined with the font attributes, e.g. “this string in this font”)
- Post-processing: none (select only), hide, hide all but selected, move to a layer (name editable, `_text` by default), or bulk edit

### Usage

1. Select the reference text, if you are searching by font attributes.
2. Run the script.
3. Set the conditions and the post-processing, then run it.

### Notes

- When moving to a layer, the lock and visibility state of existing layers is restored.
- Bulk edit replaces the contents while keeping the `characterAttributes` intact.

### Article (Japanese)

https://note.com/dtp_tranist/n/n76f1e0937088

### Update History

- v1.2.5
- v1.2.7 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.8 (2026-09-28): The button row is now built with the shared part
- v1.2.8 (2026-09-28): Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held)
- v1.2.9 (2026-09-29): Dialog opacity changed to 98%
- v1.2.10 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.2.11 (2026-09-30): Button rows with only right-side buttons are now centered
- v1.2.12 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.2.13 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.2.14 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.2.15 (2026-10-03) When several selected texts differ in font or other attributes, the attribute preview now shows “Mixed” instead of the first text's value
- v1.2.16 (2026-10-03) Select by String is now disabled when several texts are selected, so the search field no longer grows
- v1.3.0 (2026-10-03) The top of the dialog now shows the number of targets, as in “Target texts: 3”. The search field is now multiline and taller. Select by Attribute is now Select by Font Attributes, with checkboxes for font family, style, font size and fill color (combinations are matched by the script, not Illustrator's built-in commands). Removed Opacity. Select by String gained a None option and now combines with the font attributes. Option-click an attribute checkbox to check all; Option-click again to check all but that one. After Selection and the string options are laid out in two columns. None became None (select only), and the destination layer name is now editable. Text Type is now a set of checkboxes that filter and combine with the attribute and string options (all on by default; Option+T for path text). Bulk edit is available only when a Select by String option is chosen, and supports paragraph and forced line breaks via buttons and Shift+Enter. Reviewed the UI wording and tooltips, and renamed the panels to Artboard, Text Type, Font Attributes and String
