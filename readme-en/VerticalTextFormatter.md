# Apply rotated Roman characters, Tate-Chu-Yoko and number rewriting to vertical text

[![Direct](https://img.shields.io/badge/Direct%20Link-VerticalTextFormatter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/VerticalTextFormatter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VerticalTextFormatter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Applies "Rotate Standard Roman Characters in Vertical Text" and Tate-Chu-Yoko to the selected vertical text in one go.
- It can also rewrite numbers as full-width digits, kanji numerals or Japanese era years, and turn separators into marks for vertical text.
- Choose the options in a dialog box and check the result in the "Conversion Check" list before running.
- Rotation and Tate-Chu-Yoko cannot be set through the DOM, so a temporary (dynamic) action applies just those Character panel items. Font, size and other formatting stay as they are.

### Features

- Presets
  - Built-in: **Standard** / **1 full-width / 2–3 TCY / 4+ kanji** / **Kanji with 十・百・千, era years** / **Rotation only**
  - Save the current settings under a name with the save icon, and delete a saved preset with the trash icon (built-ins cannot be deleted)
- Rotation and Tate-Chu-Yoko
  - **Rotate Roman characters**: applied to all of the selected text (on by default)
  - **Tate-Chu-Yoko for numbers**: half-width numbers of the lengths checked in [2 digits] [3 digits] [4 digits] (on by default, with 2 and 3 checked)
    - **Skip numbers next to separators**: leaves "000" in "3,000" and "12" in "12.5" alone (on by default)
  - **Tate-Chu-Yoko for 2–3 capitals**: all-capital words such as "PDF" or "OK"
  - **Tate-Chu-Yoko for "!!" and "!?"**: two-character pairs of exclamation and question marks (full-width marks become half-width first)
  - **Fixed-width forms by length**: sets Tate-Chu-Yoko text to half-width forms for 2 characters, third-width for 3 and quarter-width for 4, so it fits in about one em (on by default). Nothing changes if the font has no such forms
- Character Conversion
  - **Full-width letters and digits to half-width**: makes "２０" or "ＰＤＦ" targets for Tate-Chu-Yoko and rotation
  - **Thousands "," / "，" to "、"**: 3,000 → 3、000 (on by default)
  - **Decimal "." / "．" to "・"**: 12.5 → 12・5 (on by default)
- Numbers
  - **Single digits to full-width**: keeps a single digit upright
  - **Western years to Japanese era**: 2026年 → 令和8年. When the month and day follow, they decide which side of an era change the date falls on (2019年4月30日 → 平成31年, 2019年5月1日 → 令和元年); a bare year of an era change uses the new era. Year 1 is written 元年
  - **Convert to kanji numerals**: numbers within the digit range (4–20 by default)
    - **Digit by digit**: 2026 → 二〇二六, 3,000 → 三、〇〇〇
    - **With 十・百・千**: 2026 → 二千二十六, 2000 → 二千, 3,000 → 三千 (also 万・億・兆・京; decimals stay digit by digit)
- Paragraph
  - **Mojikumi: Solid (ベタ組み)** (on by default)
  - **Kinsoku: Soft v2** (on by default)
  - **Justify with last line top** (on by default)
- **Convert horizontal text to vertical**: turns horizontal text vertical before processing (on by default)
- **Vertical text only**: horizontal text in the selection is left unchanged (on by default; disabled while "Convert horizontal text to vertical" is on)
- **Conversion Check**: lists what will change, such as "3,000 → 三千" or "10 [TCY] ×3", and updates whenever an option or digit range changes
- A progress bar shows how far processing has gone
- Settings are saved and restored the next time
- Japanese / English UI switches automatically

### How to use

1. Select vertical text objects (or a range of text with the Type tool)
2. Run the script and choose the options in the dialog box
3. Check the result in "Conversion Check" and click OK

### Order of processing

1. Turn horizontal text vertical
2. Paragraph settings (Solid mojikumi, Soft v2 kinsoku, justify with last line top)
3. Apply rotation to all of the text (rewritten characters inherit the rotated formatting)
4. Rewrite characters (full-width to half-width, separators, full-width digits, era years, kanji numerals)
5. Apply Tate-Chu-Yoko to each target

### Notes

- When the kanji and Tate-Chu-Yoko digit ranges overlap, kanji numerals win
- Without the era conversion, the four digits of a year such as 2026年 also get Tate-Chu-Yoko when [4 digits] is checked
- When a rewrite changes the number of characters (3,000 → 三千 etc.), the rewritten part takes the formatting of its first character
- Locked or hidden text is skipped

### Update History

- v1.6.0 (2026-10-15) Added presets (four built-ins, plus saving and deleting your own)
- v1.5.1 (2026-10-15) Made the dialog box more compact (shorter labels with tooltips, digit checkboxes and kanji styles on one row, a shorter check list)
- v1.5.0 (2026-10-15) Added "Convert horizontal text to vertical"
- v1.4.0 (2026-10-15) Added "Justify with last line top" to the Paragraph panel
- v1.3.0 (2026-10-15) Rotation now runs before the character rewriting (words such as "Web" or "PDF" were sometimes left unrotated). Added a progress bar
- v1.2.0 (2026-10-15) Added the Paragraph panel to apply the Solid (ベタ組み) mojikumi set and Soft v2 kinsoku
- v1.1.0 (2026-10-15) Tate-Chu-Yoko digits are now chosen with [2 digits] [3 digits] [4 digits] checkboxes. Added "Fixed-width forms by length"
- v1.0.0 (2026-10-15) Initial release
