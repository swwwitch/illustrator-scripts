# Convert full-width and half-width characters by type

[![Direct](https://img.shields.io/badge/Direct%20Link-CharacterWidthConverter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/CharacterWidthConverter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CharacterWidthConverter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README-en.md)

---

### Overview

Converts katakana, letters, digits, spaces, punctuation, brackets and symbols between half-width and full-width, choosing "Keep", "Half" or "Full" for each kind of character. It can also switch digits by length, unify Japanese punctuation and remove spaces around brackets, and presets such as vertical and Western typesetting switch all the settings at once. Character formatting (font, size, color, …) is kept character by character.

The InDesign version ([IdCharacterWidthConverter](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCharacterWidthConverter.md)) has the same settings and presets.

### Main Features

- 43 rows, each set to "Keep", "Half" or "Full", grouped into panels
  - Character Types: katakana, letters, digits, space
  - Punctuation: Japanese punctuation, middle dot, comma, period, colon, semicolon, exclamation and question marks, !? (and ?!), runs of two or more !
  - Brackets: double quote and apostrophe, parentheses, square and curly brackets, less/greater-than signs, corner brackets
  - Symbols: 17 including hyphen, underscore and the yen sign
- Half-width voiced kana (ｶﾞ, ﾊﾟ, …) become a single full-width character (ガ, パ); the reverse splits them into two characters
- "!?" and runs of "!" are handled as units; when kept, they stay as they are whatever the exclamation and question rows say
- Extra options
  - Character Types: one digit full-width and two or more half-width; convert spaces only between Japanese characters
  - Punctuation: punctuation style (unify to ，． or 、。)
  - Brackets: remove spaces inside brackets; remove spaces outside brackets; form of full-width quotes (＂＇ or “ ” ‘ ’)
  - Symbols: form of the full-width hyphen (－, ‐, −) tilde (～, 〜) and asterisk (＊, ※)
- Scope: Selection or Document
- Presets switch all the settings at once. Six are built in: "Standard", "Letters, digits & symbols to half-width", "Letters, digits & symbols to full-width", "Katakana to full-width only", "For vertical typesetting" and "For Western typesetting". "Save..." stores the current settings under a name, and "Delete" removes your own presets ("Custom" is shown when no preset matches)
- Option-click a radio button to set every row in that panel to the same option
- "Restore Defaults" brings back the initial settings (katakana to full-width, letters and digits to half-width, Japanese punctuation and corner brackets to full-width, everything else kept)
- Shows how many places were converted per row
- The settings and scope are remembered until Illustrator quits. Saved presets survive a restart

### Usage

1. Select the text objects, or select text with the Type tool (select nothing for the whole document)
2. Run the script
3. Pick a preset, or set each row to Keep, Half or Full and choose the options
4. Choose the scope and click "Convert"

### Notes

- Illustrator has no way to group a script's changes into one undo step, so undoing may take several steps.
- The steps run in this order: row conversion, punctuation style, digits by length, spaces around brackets. For example, with "," set to full-width and the punctuation style set to 、。, "," becomes "、".
- With the Selection tool, the whole text is processed, including text inside groups. With text selected by the Type tool, only the selected text is processed.
- The available scopes depend on the selection; with nothing selected only "Document" is available.
- "Document" skips locked or hidden text, including text in locked or hidden layers and groups. Threaded text is handled as one; it is skipped when any of its frames is locked or hidden.
- The katakana row also covers the prolonged sound mark (ー/ｰ) and standalone voicing marks. The middle dot (・/･) has its own row in Punctuation. Converting katakana to half-width also changes "ー" after hiragana or kanji.
- Katakana with no half-width form (ヰ, ヱ, ヮ, ヵ, ヶ) are left as they are.
- With "one digit full-width, two or more half-width", numbers such as 3.5 or 1,000 count as one number, and digits next to letters (A4) become half-width. The digits row is ignored while this is on.
- For "convert spaces only between Japanese characters", Japanese characters are kana, kanji, Japanese punctuation and full-width characters.
- "Remove spaces inside brackets" ignores straight quotes, < > and ‘ ’ (which look like apostrophes). "Remove spaces outside brackets" only applies to full-width brackets, so "word (note)" keeps its space.
- With curly full-width quotes, quotes open and close by their order in the paragraph (first opens, second closes), and a ' between letters becomes an apostrophe (’). Converting to half-width turns “ ” ‘ ’ back into " '.
- With text selected by the Type tool, characters outside the selection are not looked at: space neighbors and quote order are judged within the selection.
- With the asterisk form set to "※ (reference mark)", * and ※ are converted into each other (full-width turns * into ※, half-width turns ※ into *).
- Converting hyphens, tildes, asterisks or quotes to half-width only affects the chosen form (with "〜" chosen, "～" is not converted).
- The Japanese punctuation row converts "。"/"｡" and "、"/"､"; the corner bracket rows convert "「"/"｢" and "」"/"｣".
- The built-in presets other than "For vertical typesetting" and "For Western typesetting" convert katakana, Japanese punctuation and corner brackets to full-width and keep spaces. The two "Letters, digits & symbols" presets convert ASCII punctuation, quotes, brackets and the yen sign the same way.
- "For vertical typesetting" sets digits to "one digit full-width, two or more half-width" (for tate-chu-yoko), makes "!?" and runs of "!" half-width, spaces between Japanese characters full-width and tildes wave dashes (〜), and converts katakana, punctuation, brackets and symbols to full-width. Letters, commas, periods and quotes are kept so English words and numbers like 1,000 stay intact. It does not apply tate-chu-yoko.
- "For Western typesetting" converts full-width letters, digits, symbols, ASCII punctuation, brackets, ideographic spaces (U+3000) and the yen sign to half-width, and keeps katakana, Japanese punctuation and corner brackets.
- Presets are saved to `~/Library/Application Support/illustrator-scripts/CharacterWidthConverter.json` (under Folder.userData).

### Script info

- Version: v1.0.0

### Update History

- v1.0.0 (2026-10-30) Initial release
