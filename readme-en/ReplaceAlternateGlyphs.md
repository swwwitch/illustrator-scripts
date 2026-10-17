# Switch glyph forms and variants of the selected characters

[![Direct](https://img.shields.io/badge/Direct%20Link-ReplaceAlternateGlyphs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceAlternateGlyphs.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceAlternateGlyphs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Switches the glyph form of the selected characters (JIS78, JIS90, Traditional, etc.) from a dialog box, applied on the spot.
- Also replaces characters with variant characters (such as 高 → 髙) and lays out a list of the variants with their code points.

### Features

- The top of the dialog box shows the selected characters with their code points, as in "葛 U+845B" (up to 5; updated after a variant replacement)
- Glyph Form (left)
  - Default Form, JIS78, JIS83, JIS90, JIS04, Traditional, Expert, Full Widths, Half Widths, Third Widths, Quarter Widths
  - The form shared by the selected characters is checked when the dialog opens. Clicking one applies it right away
  - Forms the font lacks (those that look like the default for every selected character) are dimmed. Up to the first 10 characters are checked
- Variant Characters (right)
  - When the selection contains characters with registered variants, the candidates are listed as "髙  U+9AD9". Clicking one replaces every occurrence of that character in the selection right away
  - With several such characters, each line carries the original, as in "高 → 髙  U+9AD9"
  - The original character is among the candidates, so choosing it again restores it
- Create List
  - Closes the dialog box and makes a list on the layer "// glyphs" (created on top if missing): two rows (Horizontal and Vertical) per selected character, showing only the distinct shapes among the 11 glyph forms and the registered variants in large type with code points and form names (forms the font lacks look like the default and are left out). The 11 forms plus Proportional Widths are tried on the registered variants too, and the vertical-setting forms are listed as well (such as the vertical prolonged sound mark). It then tries the source character and each variant with variation selectors (U+E0100 onward; IVS are registered per code point, so starting from 辺 still reaches the IVS of 邉) and adds the IVS glyphs, which covers many of the alternates in the Glyphs panel (the note shows the selector's code point), to the right of the source text. Sample glyphs inherit the source character's font and size; code points and form names use Hiragino Kaku Gothic W3
  - Puts a black rectangle (30% opacity) behind the list
  - Selects the list and zooms so it fills about 65% of the window
  - Glyph form and variant changes made so far are kept
- List panel (above the buttons)
  - Create on the [// glyphs] layer: adds the list to the same layer, keeping earlier lists
  - Delete the [// glyphs] layer, then create (default): removes the previous list layer first; skipped when the selected text is on that layer
  - Create List button
- Cancel restores the glyph forms and variants applied along the way
- Japanese / English UI switches automatically

### How to use

1. Select characters with the Type tool (or select text objects)
2. Run the script
3. Click a glyph form or a variant candidate
4. Click OK

### Options

Edit the User Settings at the top of the script.

- `ITAIJI_GROUPS`: variant groups. Each string lists characters that are variants of one another (e.g. `"辺邊邉"`). Add lines to extend it; groups sharing a character are merged at run time
- `GLYPH_FORM_CHECK_MAX`: number of characters checked for supported forms (default 10; more makes opening slower)
- `GLYPH_LIST_LAYER_NAME`: the layer for the list (default `// glyphs`)
- `GLYPH_LIST_CODE_RATIO`: code point size relative to the sample glyph (default 0.25; form names are 3/4 of it)
- `GLYPH_LIST_LABEL_MIN_SIZE`: minimum size of code points and form names (pt, default 6)
- `GLYPH_LIST_LABEL_FONT_NAMES`: font for code points and form names (PostScript names, tried in order; falls back to the source font)
- `GLYPH_LIST_IVS_MAX` / `GLYPH_LIST_IVS_MISS_LIMIT`: number of variation selectors tried (default 64) and how many in a row without a new shape end the search (default 3)
- `DELETE_LIST_LAYER_DEFAULT`: true opens with "Delete the [// glyphs] layer, then create" chosen, false with "Create on the [// glyphs] layer" (default true)
- `GLYPH_LIST_BACK_OPACITY` / `GLYPH_LIST_BACK_PADDING`: opacity of the black backdrop (%, default 30) and its padding relative to the first sample size (default 1)
- `GLYPH_LIST_COLUMNS`: samples per line before wrapping (default 5)
- `GLYPH_LIST_FIT_PERCENT`: share of the window the list fills (%, default 65)

### Notes

- Supported forms are checked on opening by outlining the characters on a temporary layer
- Each character gets two rows, Horizontal and Vertical, with a heading on the left; samples wrap every 5. The Vertical row holds only shapes that differ from horizontal setting, as vertical point text (omitted when there are none)
- IVS samples appear only when the font supports IVS. Alternates without an IVS assignment (shown only in the Glyphs panel) cannot be reached from a script and are not listed
- Create List closes the dialog box; run the script again to keep switching
- Variants that are compatibility ideographs (e.g. U+FA45) may not display in some fonts
- Spaces and line breaks are skipped. Threaded text is handled as one story; locked or hidden text is skipped

### Update History

- v1.1.1 (2026-10-17) Added space below "Selected" at the top of the dialog box
- v1.1.0 (2026-10-17) Create List improved: distinct shapes among the glyph forms, variants, IVS and vertical forms are laid out in Horizontal and Vertical rows, five per line, with code points and form names (Hiragino Kaku Gothic W3). Samples inherit the source font and size, sit on a 30% black backdrop, and are selected and zoomed to about 65% of the window. Added the List panel (create on the layer / delete it first). Forms the font lacks are dimmed, and the selected characters' code points appear at the top of the dialog box
- v1.0.0 (2026-10-17) Initial release
