# Create a composite font from Japanese, Kana and Roman fonts

[![Direct](https://img.shields.io/badge/Direct%20Link-CompositeFontMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/CompositeFontMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CompositeFontMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Creates an Illustrator composite font from just three fonts: Japanese, Kana and Roman. Illustrator has no scripting command for composite fonts, so the script writes a file in the same format the Composite Fonts dialog saves, straight into the composite font folder. Restart Illustrator to use the new composite font.

### Main Features

- Only three fonts to choose
  - **Japanese**: Kanji, full-width punctuation, full-width symbols
  - **Kana**: hiragana, katakana
  - **Roman**: half-width alphabetic characters and numerals
- Set the size and baseline of Kana and Roman relative to the Japanese font (Kanji)
- Picks up fonts from the selected text as the initial values
- When Kana or Roman in the selected text differs in size from the Japanese text, uses the ratio as the initial size (e.g. 10 pt Japanese and 10.8 pt Roman → 108%)
- The composite font name has two fields joined with a hyphen: [JapanesePS-KanaPS-RomanPS]-[weight]
  - The name part fills in from the chosen fonts' PostScript names with the weight removed (-Bold, -W3, ...; e.g. PA1MinchoStdN-Bold → PA1MinchoStdN); Kana is omitted when same as Japanese
  - The weight fills in from the Japanese style name (spaces and non-ASCII removed; left empty, no weight is added)
- Custom sets: load the custom sets from an existing composite font (such as sw-B) and set the font, size and baseline of each
- Numeric fields step to the next integer with the ∧∨ buttons and the arrow keys (1.5 → 2); add Shift for the next multiple of 10, Option for ±0.1

### How to Use

1. Select the text to use as the font sample and run the script (it also runs with nothing selected)
2. Check the Japanese, Kana and Roman fonts (family and style)
3. Adjust the size and baseline of Kana and Roman if needed
4. To use custom sets, choose the source composite font under Load from in the Custom Character Sets panel, then adjust each set
5. Check the composite font name (name part and weight) and click Create
6. Restart Illustrator; the composite font appears in the font menu

### Reading the Selected Text

The selected texts are ordered top to bottom, and the font of each one's first character is picked up.

| Texts from the top | Assigned to |
|---|---|
| 1 | Japanese |
| 2 | Japanese, Roman (Kana uses the Japanese font at 100% / 0%, and its row is dimmed) |
| 3 | Japanese, Kana, Roman |

- With several text objects selected, they are ordered by their top edge (left to right at the same height)
- With one text object, or characters selected, paragraphs are used in order
- Anything not picked up falls back to Kozuka Gothic Pr6N R (Japanese) and Myriad Pro Regular (Roman)

### Custom Character Sets

Choose a composite font that has custom sets under Load from in the Custom Character Sets panel, and its sets (parentheses, small kana, prolonged sound mark, and so on) appear one per row.

- Only composite fonts in the composite font folder that have custom sets are listed
- Each set's initial font follows its characters: Roman when every character is half-width alphabetic or numeric, Kana when all are kana, otherwise Japanese. Size and baseline start from that group's values too
- Hover over a set name to see its characters
- Changing the source reopens the dialog, keeping your entries
- The Chars… button on each row shows and edits the set's characters, one per character; duplicates, spaces and line breaks are ignored
- Custom sets cannot be created in the dialog; use a composite font whose sets were made in Illustrator's Custom Character Sets as the source

### Notes

- Composite font names (including the hyphen and weight) are limited to 29 ASCII letters, digits and symbols (no "/" or ":"). A composite font with a 43-character name kept Illustrator from launching; names up to 29 characters are confirmed to work
- The automatic name part is cut so that it fits in 29 characters together with the weight
- If a composite font with the same name exists, the script asks before overwriting it
- The Japanese font stays at 100% size and 0% baseline; vertical and horizontal scale are always 100%
- Variable fonts are not listed. A composite font containing a variable font crashed Illustrator when applied. Fonts whose PostScript or family name contains "VF" or "Var" are not listed; variable fonts that the name does not reveal (such as Bodoni Moda) are caught on Create, with a warning. Variable fonts in the selected text are not used as initial values either
- Files are saved to `~/Library/Application Support/Adobe/Adobe Illustrator <version>/ja_JP/合成フォント/`; if that folder is not found, you choose one
- If Illustrator ever fails to launch, move the composite font file you created out of that folder

### Update History

- v1.0.0 (20260927): Initial release
