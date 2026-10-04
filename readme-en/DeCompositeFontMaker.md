# Replace composite fonts with their component fonts

[![Direct](https://img.shields.io/badge/Direct%20Link-DeCompositeFontMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/DeCompositeFontMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeCompositeFontMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Replaces the composite fonts in the selected text with their component fonts (the Japanese, Kana, Roman and other fonts), character by character. The size, scale and baseline set in the composite font are carried over to the character formatting, so the text keeps its look with ordinary fonts. Useful when handing files to someone who does not have the composite font.

Illustrator scripting cannot read what is inside a composite font, so the script reads the file in the composite font folder directly.

### Main Features

- Works out which character set each character belongs to (Kanji, Kana, full-width punctuation, full-width symbols, half-width Roman, half-width numerals) and applies that set's font
- Carries the composite font's settings over to the character formatting
  - **Size**: multiplied into the font size (e.g. 24 pt with Roman at 116% → 27.84 pt)
  - **Vertical / horizontal scale**: multiplied into the character's scale
  - **Baseline**: added to the baseline shift, relative to the original size (e.g. 24 pt at −2% → −0.48 pt)
- Characters with auto leading keep their current leading as a fixed value (when the composite font has a resized set)
- Tracking and manual kerning are rescaled to the new size (both are in 1/1000 em)
- Only characters set in a composite font are replaced; the rest stay as they are
- An alert reports how many characters were replaced and why any were not

### How to Use

1. Select text objects (text inside groups, or selected characters, also work)
2. Run the script

### When Characters Are Not Replaced

| Situation | Result |
|---|---|
| The composite font file is not on this machine | Those characters stay as they are; the component font names (file names) recorded in the document are reported |
| A component font is not installed | Those characters stay as they are; the font name and character count are reported |

### Notes

- Custom character sets are ignored. Their characters are replaced by the settings of the standard sets (Kanji, Kana, Roman, ...)
- Composite font files are looked up in `~/Library/Application Support/Adobe/Adobe Illustrator <version>/ja_JP/合成フォント/`
- Pinning the leading and rescaling tracking / kerning can be turned off in the user settings at the top of the script (`KEEP_LEADING` / `KEEP_SPACING`)

### Update History

- v1.0.0 (20260927): Initial release
- v1.0.1 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
