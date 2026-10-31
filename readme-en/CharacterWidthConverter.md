# Convert full-width and half-width characters by type

[![Direct](https://img.shields.io/badge/Direct%20Link-CharacterWidthConverter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/CharacterWidthConverter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CharacterWidthConverter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README-en.md)

---

### Overview

Converts katakana, letters, digits, spaces, punctuation, brackets and symbols between half-width and full-width, choosing "Keep", "Half" or "Full" for each kind of character. It can also switch digits by length, unify Japanese punctuation and remove spaces around brackets, and presets such as vertical and Western typesetting switch all the settings at once. The Horizontal and Vertical tabs each keep their own settings, and the open tab's settings are used. Character formatting (font, size, color, …) is kept character by character.

The InDesign version ([IdCharacterWidthConverter](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCharacterWidthConverter.md)) has the same settings and presets.

### Main Features

- The Horizontal and Vertical tabs each keep their own settings, and the open tab's settings are used. Options only for vertical text are on the Vertical tab, options for Western text on the Horizontal tab
- 43 rows, each set to "Keep", "Half" or "Full", grouped into panels
  - Character Types: katakana, letters, digits, space
  - Punctuation: Japanese punctuation, comma, period, colon, semicolon, exclamation and question marks, !? (and ?!), runs of two or more !
  - Brackets: double quote and apostrophe, parentheses, square and curly brackets, less/greater-than signs, corner brackets
  - Symbols: 18 including hyphen, underscore, the yen sign and the middle dot
- Half-width voiced kana (ｶﾞ, ﾊﾟ, …) become a single full-width character (ガ, パ); the reverse splits them into two characters
- "!?" and runs of "!" are handled as units; when kept, they stay as they are whatever the exclamation and question rows say
- Extra options
  - Character Types: one digit full-width and two or more half-width; convert spaces only between Japanese characters; remove spaces between Japanese and Western text; add a space after , . : in Western text (Horizontal tab; apples,oranges → apples, oranges)
  - Punctuation: punctuation style (unify to ，． or 、。); convert to ... (・・・, ... and 。。。 become an ellipsis …)
  - Brackets: remove spaces inside brackets; remove spaces outside brackets; form of full-width quotes (＂＇, “ ” ‘ ’ or 〝 〟)
  - Symbols: form of the full-width hyphen (－, ‐, −) tilde (～, 〜) and asterisk (＊, ※)
- Scope: Selection or Document
- Vertical Text (Vertical tab): numbers of 3+ digits to full-width (2026 → ２０２６); "," in numbers to "、" (1,000 → 1、000); "." in numbers to "・" (12.5 → 12・5); all-capital words of up to 3 letters to full-width (PDF → ＰＤＦ)
- Check Coloring fills converted characters with red (removed spaces are not marked)
- Presets switch all the settings at once. Six are built in: "Standard", "Letters, digits & symbols to half-width", "Letters, digits & symbols to full-width", "Katakana to full-width only", "For vertical typesetting" and "For Western typesetting". Presets apply to the open tab, except that "For vertical typesetting" switches to the Vertical tab and "For Western typesetting" to the Horizontal tab. "Save..." stores the current settings under a name, and "Delete" removes your own presets (choosing a saved preset switches to the tab it was saved from; "Custom" is shown when no preset matches)
- Option-click a radio button to set every row in that panel to the same option
- "Restore Defaults" brings back the initial settings of the open tab (katakana to full-width, letters and digits to half-width, Japanese punctuation and corner brackets to full-width, everything else kept)
- Shows how many places were converted per row
- The settings of each tab, the open tab and the scope are remembered until Illustrator quits. Saved presets survive a restart

### Usage

1. Select the text objects, or select text with the Type tool (select nothing for the whole document)
2. Run the script
3. Open the Horizontal or Vertical tab, then pick a preset, or set each row to Keep, Half or Full and choose the options
4. Choose the scope and click "Convert"

### Notes

- Illustrator has no way to group a script's changes into one undo step, so undoing may take several steps.
- Only the open tab's settings are used. The text direction is not checked, so vertical text can be converted with the Horizontal tab's settings too.
- Presets saved with v1.2.0 or earlier apply to the open tab (options not on the tab are ignored).
- The steps run in this order: ellipses (・・・, ..., 。。。 → …), spaces between Japanese and Western text, row conversion, punctuation style, digits by length (one, two or more, three or more), capital words of up to 3 letters, "," and "." in numbers, spaces after , . : in Western text, spaces around brackets. For example, with "," set to full-width and the punctuation style set to 、。, "," becomes "、".
- With the Selection tool, the whole text is processed, including text inside groups. With text selected by the Type tool, only the selected text is processed.
- The available scopes depend on the selection; with nothing selected only "Document" is available.
- "Document" skips locked or hidden text, including text in locked or hidden layers and groups. Threaded text is handled as one; it is skipped when any of its frames is locked or hidden.
- The katakana row also covers the prolonged sound mark (ー/ｰ) and standalone voicing marks. The middle dot (・/･) has its own row in Symbols. Converting katakana to half-width also changes "ー" after hiragana or kanji.
- Katakana with no half-width form (ヰ, ヱ, ヮ, ヵ, ヶ) are left as they are.
- With "one digit full-width, two or more half-width", numbers such as 3.5 or 1,000 count as one number, and digits next to letters (A4) become half-width. The digits row is ignored while this is on.
- "Numbers of 3+ digits to full-width" also counts numbers with separators as one (1,000 → １,０００; "," and "." stay as they are) and leaves digits next to letters alone. Numbers of one or two digits follow the digits row (or "one digit full-width, two or more half-width"), so with both on, one digit is full-width, two are half-width and three or more full-width.
- "All-capital words of up to 3 letters to full-width" applies to words of one to three capital letters, half- or full-width (A → Ａ, NHK → ＮＨＫ). Words with lowercase letters, digits or marks (Pc, A4, U.S., AT&T) and longer words are left alone. These words become full-width even when the letters row is set to Half.
- For "convert spaces only between Japanese characters", Japanese characters are kana, kanji, Japanese punctuation and full-width characters.
- "Remove spaces inside brackets" ignores straight quotes, < > and ‘ ’ (which look like apostrophes). "Remove spaces outside brackets" only applies to full-width brackets, so "word (note)" keeps its space.
- With curly full-width quotes, quotes open and close by their order in the paragraph (first opens, second closes), and a ' between letters becomes an apostrophe (’). Converting to half-width turns “ ” ‘ ’ back into " '.
- With text selected by the Type tool, characters outside the selection are not looked at: space neighbors and quote order are judged within the selection.
- With the asterisk form set to "※ (reference mark)", * and ※ are converted into each other (full-width turns * into ※, half-width turns ※ into *).
- Converting hyphens, tildes, asterisks or quotes to half-width only affects the chosen form (with "〜" chosen, "～" is not converted).
- The Japanese punctuation row converts "。"/"｡" and "、"/"､"; the corner bracket rows convert "「"/"｢" and "」"/"｣".
- The built-in presets other than "For vertical typesetting" and "For Western typesetting" convert katakana, Japanese punctuation and corner brackets to full-width and keep spaces. The two "Letters, digits & symbols" presets convert ASCII punctuation, quotes, brackets and the yen sign the same way.
- "For vertical typesetting" sets digits to "one digit full-width, two or more half-width" (for tate-chu-yoko), makes "!?" and runs of "!" half-width, spaces between Japanese characters full-width and tildes wave dashes (〜), and converts katakana, punctuation, brackets and symbols to full-width. The letter, comma, period and quote rows are kept so English words stay intact, while "," and "." in numbers become "、" and "・" (1,000 → 1、000, 12.5 → 12・5). Spaces between Japanese and Western text are removed, and ・・・, ... and 。。。 become …. It does not apply tate-chu-yoko.
- "For Western typesetting" converts full-width letters, digits, symbols, ASCII punctuation, brackets, ideographic spaces (U+3000) and the yen sign to half-width, and keeps katakana, Japanese punctuation and corner brackets. It turns ... into … and adds a space after , . : in Western text.
- Presets are saved to `~/Library/Application Support/illustrator-scripts/CharacterWidthConverter.json` (under Folder.userData).

### Article

https://note.com/dtp_tranist/n/nf9d7245b0653

### Script info

- Version: v1.3.0

### Update History

- v1.3.0 (2026-11-01) Split the settings into Horizontal and Vertical tabs, each with its own settings; the open tab's settings are used. Vertical Text is on the Vertical tab only, "Add a space after , . : in Western text" on the Horizontal tab only. Added "Numbers of 3+ digits to full-width" and "All-capital words of up to 3 letters to full-width" to Vertical Text. Moved Check Coloring next to Scope. Presets apply to the open tab; choosing "For vertical typesetting" switches to the Vertical tab and "For Western typesetting" to the Horizontal tab. Choosing a saved preset switches to the tab it was saved from
- v1.2.0 (2026-10-31) Added "Remove spaces between Japanese and Western text" (第 3 章 → 第3章) and "Add a space after , . : in Western text" (1,000, example.com and 10:30 are left alone) to Character Types. Added "Convert to ..." (・・・, ..., 。。。) to Punctuation: three middle dots, periods or Japanese full stops in a row, half- or full-width, become an ellipsis (…). The "For vertical typesetting" preset now removes spaces between Japanese and Western text and converts to ellipses; "For Western typesetting" turns ... into … and adds spaces after , . :. Renamed the dialog to "Convert Width & Clean Up Text"
- v1.1.0 (2026-10-30) Moved the middle dot row from Punctuation to Symbols. The dialog now has three columns, with Scope and Presets on one row, so it fits the screen height. The full-width form menus (hyphen, tilde, asterisk, quotes) are dimmed while their row is set to Keep. Added minutes (〝 〟) as a full-width form of double quotes. Added Vertical Text options that turn "," in numbers into "、" and "." into "・" (on in the "For vertical typesetting" preset). Added Check Coloring to color converted characters
- v1.0.0 (2026-10-30) Initial release
