# Switch glyph forms and variants of the selected characters

[![Direct](https://img.shields.io/badge/Direct%20Link-ReplaceAlternateGlyphs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceAlternateGlyphs.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceAlternateGlyphs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Switches the glyph form of the selected characters (JIS78, JIS90, Traditional, etc.) from a dialog box, applied on the spot.
- Also replaces characters with variant characters (such as 高 → 髙) and lays out a list of the variants with their code points.

### Features

- Glyph Form (left)
  - Default Form, JIS78, JIS83, JIS90, JIS04, Traditional, Expert, Full Widths, Half Widths, Third Widths, Quarter Widths
  - The form shared by the selected characters is checked when the dialog opens. Clicking one applies it right away
- Variant Characters (right)
  - When the selection contains characters with registered variants, the candidates are listed as "髙  U+9AD9". Clicking one replaces every occurrence of that character in the selection right away
  - With several such characters, each line carries the original, as in "高 → 髙  U+9AD9"
  - The original character is among the candidates, so choosing it again restores it
- Create List (lower left)
  - Lays out the variants in large type with their code points on the layer "_glyph" (created on top if missing), to the right of the source text, in the source character's font
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
- `GLYPH_LIST_LAYER_NAME`: the layer for the list (default `_glyph`)
- `GLYPH_LIST_GLYPH_SIZE` / `GLYPH_LIST_CODE_SIZE`: sizes of the glyphs and code points in the list (pt)

### Notes

- A glyph form changes nothing visible when the font lacks it
- A list made with Create List stays even if you cancel
- Variants that are compatibility ideographs (e.g. U+FA45) may not display in some fonts
- Spaces and line breaks are skipped. Threaded text is handled as one story; locked or hidden text is skipped

### Update History

- v1.0.0 (2026-10-17) Initial release
