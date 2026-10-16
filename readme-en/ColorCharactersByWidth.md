# Color full-width and half-width alphanumerics and fixed-width glyphs differently

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorCharactersByWidth.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ColorCharactersByWidth.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorCharactersByWidth.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Colors full-width alphanumerics and "！？" red, half-width alphanumerics and "!?" blue, and characters set to fixed-width glyphs green in the selected text. All other characters get the base color (black).
- Half-width and full-width spaces cannot be colored, so colored rectangles are drawn on a dedicated layer.
- Useful for spotting mixed full-width and half-width characters and fixed-width glyph settings at a glance.

### Main Features

- Target characters
  - Full-width: ０–９, Ａ–Ｚ, ａ–ｚ, ！, ？
  - Half-width: 0–9, A–Z, a–z, !, ?
  - Fixed-width glyphs: characters set to the OpenType full-, half-, third- or quarter-width glyphs (not only alphanumerics)
- Fixed-width glyphs take priority over full-width and half-width.
- Characters that match none of these get the base color.
- Spaces (black for both half-width and full-width by default) get rectangles at 50% opacity on the frontmost layer "ColorNumberByWidth Spaces". Each run clears the rectangles on this layer and draws them again.
- Turn off Draw Spaces to skip the space rectangles and leave the "ColorNumberByWidth Spaces" layer untouched (earlier rectangles stay). While it is off, the other settings in the Spaces panel are dimmed (on by default).
- Turn on Group Space Rectangles to put the rectangles into one group (on by default).
- Turn off a setting's checkbox to leave that kind uncolored (no rectangles for spaces). With Base off, characters that match nothing keep their color. With Fixed-width Glyphs off, those characters are judged as full- or half-width.
- Click a swatch in the dialog to choose the color with Illustrator's standard color picker.
- Colors are created in the document's color mode (RGB / CMYK).

### Usage

1. Select the text (a range selected with the Type tool also works)
2. Run the script
3. Click the Base, Full-width, Half-width or Fixed-width Glyphs, Half-width Space or Full-width Space swatch to change the color if needed
4. Click OK

### Notes

- With the Selection tool, the whole text is processed, including text inside groups.
- Other symbols (＃, ％, #, etc.) get the base color.
- Space rectangles are drawn for both horizontal and vertical text, but not for threaded text. They are misplaced on rotated or sheared text.
- A rectangle is as tall as the type size, with its bottom at the underline position. It may sit slightly high or low depending on the font. In vertical text it is as wide as the type size, aligned to the right edge of the em box.
- Change the initial colors with `FULL_WIDTH_RGB` and the other values at the top of the script. Chosen colors are not remembered for the next run.

### Script info

- Version: v1.4.0

### Update History

- v1.4.0 (2026-10-16) Renamed from ColorNumberByWidth. Added Draw Spaces to the Spaces panel. When off, no space rectangles are drawn and the panel's other settings are dimmed
- v1.3.0 (2026-10-16) Added colored rectangles over half-width and full-width spaces, drawn on a dedicated layer, with an option to group them. Each setting has a checkbox; unchecked kinds are not colored or drawn
- v1.2.1 (2026-10-16) The button row is back to the shared part's standard layout (centered when the dialog's inner width is up to 240 px)
- v1.2.0 (2026-10-16) Added a base color (black by default) for characters that match none of the others. The buttons are now centered
- v1.1.0 (2026-10-16) Added coloring characters set to fixed-width glyphs (full-, half-, third- and quarter-width) green. Changed the default half-width color to #0099ff (C100 M40 in CMYK)
- v1.0.0 (2026-10-16) Initial release
