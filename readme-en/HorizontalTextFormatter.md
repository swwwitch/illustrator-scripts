# Turn text formatted for vertical setting back into horizontal text

[![Direct](https://img.shields.io/badge/Direct%20Link-HorizontalTextFormatter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/HorizontalTextFormatter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/HorizontalTextFormatter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- The reverse of [VerticalTextFormatter](VerticalTextFormatter.md).
- Turns vertical text horizontal, removes Roman rotation and Tate-Chu-Yoko, and rewrites characters changed for vertical setting back to horizontal style.
- Choose the options in a dialog box and check the result in the "Conversion Check" list before running.
- Roman rotation and Tate-Chu-Yoko cannot be set through the DOM, so a temporary (dynamic) action turns off just those Character panel items.

### Features

- Orientation and Rotation
  - **Make vertical text horizontal** (every frame of threaded text; on by default)
  - **Remove Roman rotation** for the whole text (on by default)
  - **Remove Tate-Chu-Yoko** for the whole text (on by default)
  - **Reset fixed-width glyphs**: resets glyph forms such as the half-, third- and quarter-width forms used in Tate-Chu-Yoko (on by default)
- Character Conversion
  - **Full-width alnum to half-width**: ＰＤＦ → PDF, ５ → 5, ！？ too (on by default)
  - **Separator 、 → ,**: 3、000 → 3,000 (on by default)
  - **Decimal ・ → .**: 12・5 → 12.5 (on by default)
  - **Symbols to half-width**: ％ ＆ ＃ ＠ ＊ ＋ ＝ ＜ ＞ ／; brackets, ： and ； stay full-width (on by default)
  - **Quotes 〝〟 → “”** (on by default)
  - **Number hyphens to -**: ‐ ー － in phone and postal numbers become - (on by default)
  - **Unit squares to letters**: ㍉ ㌢ ㌖ ㍍ ㍉㌘ ㌘ ㌕ ㌧ ㍉㍑ ㌥㍑ ㌔㍑ ㍑ ㌍ ㌔㌍ ㏄ → mm cm km m mg g kg t ml dL kL L cal kcal cc, ㎟ ㎠ ㎡ ㎢ ㎣ ㎤ ㎥ ㎦ → mm² cm² m² km² mm³ cm³ m³ km³ (on by default)
  - **Range 〜 → -**: 10〜20 → 10-20 (off by default)
  - **Japanese era to Western year**: 令和8年 → 2026年, 令和元年 → 2019年 (off by default)
  - **Kanji numerals to Arabic**: runs of two or more kanji numerals (二〇二六 → 2026, 二千二十六 → 2026, 二万五〇〇〇 → 25000, 三、〇〇〇 → 3,000, 一二・五 → 12.5); a single one only before 月 日 時 分 秒 (2019年四月30日 → 2019年4月30日, 午後三時 → 午後3時). 一緒, 十分, 数十万, 何百 and the like are left (off by default)
- **Conversion Check**: lists what will change and updates whenever an option changes
- A progress bar shows how far processing has gone
- Settings are saved and restored the next time
- Japanese / English UI switches automatically

### How to use

1. Select vertical text objects (or a range of text with the Type tool)
2. Run the script and choose the options in the dialog box
3. Check the result in "Conversion Check" and click OK

### Order of processing

1. Make vertical text horizontal
2. Remove Roman rotation and Tate-Chu-Yoko from the whole text
3. Rewrite characters
4. Reset fixed-width glyphs

### Notes

- "Kanji numerals to Arabic" may also change words that contain two or more kanji numerals (such as 万一); check the list first
- "Reset fixed-width glyphs" resets every glyph form in the text, including forms such as JIS78
- When a rewrite changes the number of characters (令和8年 → 2026年 etc.), the rewritten part takes the formatting of its first character
- Locked or hidden text is skipped

### Update History

- v1.2.0 (2026-10-23) "Kanji numerals to Arabic" now also converts a single kanji numeral before 月 日 時 分 秒 (2019年四月30日 → 2019年4月30日), and leaves approximate numbers such as 数十万 and 何百
- v1.1.0 (2026-10-16) "Unit squares to letters" now also covers ㌥㍑ ㌔㍑ ㌍ ㌔㌍ ㏄ and ㎟ ㎠ ㎡ ㎢ ㎣ ㎤ ㎥ ㎦ (to match VerticalTextFormatter v1.22.0)
- v1.0.0 (2026-10-16) Initial release
