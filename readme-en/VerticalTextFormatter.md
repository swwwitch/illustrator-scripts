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

- Simple and Detailed modes (switch at the left of the button row; Simple by default)
  - **Simple**: shows only rotation, Tate-Chu-Yoko digits (2/3/4), fixed-width forms, separator conversion, and full-width single digits and letters
  - **Detailed**: shows every option
  - Simple mode applies only the visible options (hidden options are off; kinsoku and mojikumi stay unchanged). Detailed-mode settings are kept
  - Switching reopens the dialog box around the same center
- Presets
  - Built-in (based on common rules for numbers in vertical text)
    - **Standard**: TCY for 2–3 digits (fixed-width forms), 、 for separators, ・ for decimals
    - **Arabic, TCY for 2 digits only**: 2 digits in TCY, single digits full-width, 3+ digits upright one by one
    - **Arabic, TCY for 2–3 digits (magazines, ads)**: single digits full-width, unit squares (10cm → 10㌢), long Roman text sideways, short words full-width (TCY for numbers only)
    - **Kanji, 万・億・兆 only (newspapers, books)**: 25000 → 二万五〇〇〇, separators removed
    - **Kanji with 十・百・千 (literature)**: 25000 → 二万五千
    - **Kanji digit by digit with 、**: 3,000 → 三、〇〇〇
    - **Rotation only**
    - In the kanji presets too, years stay digit by digit (二〇二六年) and phone / postal numbers stay Arabic
  - Save the current settings under a name with the save icon, and delete a saved preset with the trash icon (built-ins cannot be deleted)
- Rotation and Tate-Chu-Yoko
  - **Rotate Roman characters**: applied to all of the selected text (on by default). Runs of two or more words joined by spaces or & (Type & Layout, Adobe Illustrator, or longer) always stay sideways and are not made full-width or changed. Full-width alphanumerics, spaces and ＆ inside such runs count too and become half-width (Design Ｗｅｅｋ → Design Week). The "magazines, ads" preset now turns on "Units to squares". Added "Height to a multiple of the type size" (on by default, always applied in Simple mode too)
    - **Skip long Roman text**: Roman text of at least the given number of characters (5 by default) stays sideways. Words joined by spaces or . ' & - count as one run ("Adobe Illustrator" is 17)
  - **Tate-Chu-Yoko for numbers**: half-width numbers of the lengths checked in [2 digits] [3 digits] [4 digits] (on by default, with 2 and 3 checked)
    - **Skip numbers next to separators**: leaves "000" in "3,000" and "12" in "12.5" alone (on by default)
  - **Tate-Chu-Yoko for 2–3 capitals**: all-capital words such as "PDF" or "OK"
  - **Tate-Chu-Yoko for "!!" and "!?"**: two-character pairs of exclamation and question marks (full-width marks become half-width first)
  - **Fixed-width forms by length**: sets Tate-Chu-Yoko text to half-width forms for 2 characters, third-width for 3 and quarter-width for 4, so it fits in about one em (on by default). Nothing changes if the font has no such forms
- Character Conversion
  - **Full-width letters and digits to half-width**: makes "２０" or "ＰＤＦ" targets for Tate-Chu-Yoko and rotation
  - **Separators**: "," "，" become 、 (3,000 → 3、000, default) / are removed (3,000 → 3000; the joined digits count as one number) / stay
  - **Decimal "." / "．" to "・"**: 12.5 → 12・5 (on by default)
  - **Range "-" "~" → "〜"**: replaces - ~ － ～ between digits with 〜 (10-20 → 10〜20, on by default). Phone and postal numbers are excluded
  - **Symbols to full-width**: makes half-width ( ) [ ] { } : ; / % & # @ * + = < > full-width so they stand upright (on by default)
  - **Quotes to 〝〟**: replaces “ ” and straight " " with 〝 〟 (on by default)
  - **"!" "?" to full-width**: ! and ? that are not in Tate-Chu-Yoko become full-width (on by default)
  - **Short words to full-width** (Detailed mode): words of at most the given number of letters become full-width (OK → ＯＫ). The length is linked to one below "Skip long Roman text" (5 and 4 by default). Unit squares and capital Tate-Chu-Yoko take priority
  - **Single letters to full-width**: a lone letter such as the A in Aランク
- Numbers
  - **Single digits to full-width**: keeps a single digit upright
  - **Units to squares**: turns mm, cm, km, m, mg, g, kg, t, ml, L right after digits into ㍉ ㌢ ㌖ ㍍ ㍉㌘ ㌘ ㌕ ㌧ ㍉㍑ ㍑ (10cm → 10㌢). Off by default because some fonts lack these glyphs
  - **Western years to Japanese era**: 2026年 → 令和8年. When the month and day follow, they decide which side of an era change the date falls on (2019年4月30日 → 平成31年, 2019年5月1日 → 令和元年); a bare year of an era change uses the new era. Year 1 is written 元年
  - **Convert to kanji numerals**: numbers within the digit range (4–20 by default)
    - **Digit by digit**: 2026 → 二〇二六, 3,000 → 三、〇〇〇
    - **With 十・百・千**: 2000 → 二千, 3,000 → 三千, 1234 → 千二百三十四 (also 万・億・兆・京; decimals stay digit by digit)
    - **Only 万・億・兆**: 25000 → 二万五〇〇〇, 123456789 → 一億二三四五万六七八九
    - A year followed by 年 is always digit by digit (2026年 → 二〇二六年)
    - **0 as 零**: uses 零 instead of 〇 (二〇〇六 → 二零零六)
    - **千 as 一千**: with 十・百・千, writes 1000 as 一千 (1500 → 一千五百)
- **Number hyphens** (Detailed mode): hyphens in phone and postal numbers (digits-digits-digits, 3-4 digits) become ‐ (hyphen, default) / ー (long vowel mark) / stay. These numbers are never ranges, Tate-Chu-Yoko or kanji
- Character and Paragraph (Detailed mode)
  - Choose the kinsoku set (None / Hard / Soft / Soft v2) and the mojikumi set (None, the seven built-ins and the document's custom sets) from dropdowns (Soft v2 and Solid by default)
  - **Leading**: set as an auto-leading percentage (175% by default, 50–500%). In both modes, leading is always set to Auto at the end so it follows the type size
  - Justification is always justified with the last line at the top
  - **Proportional metrics for full-width Latin**: applies the OpenType Proportional Metrics feature to full-width letters (after the rewriting, on by default)
  - **Height to a multiple of the type size**: sets the area-type frame height (the line length in vertical text) to the nearest whole multiple of the type size, inset excluded (on by default; always applied, in Simple mode too)
  - **Auto-size area type**: after processing, turns on Auto Size in Area Type Options so the frame fits the text (on by default and in every built-in preset; threaded text is skipped)
- Check Coloring (Detailed mode)
  - **Color characters after processing**: full-width alphanumerics and ！？ red, half-width blue, fixed-width glyphs (such as third- and quarter-width TCY) green, and the rest in the base color (black). Off by default. Click a swatch to change the color with the color picker. For checking only; the original colors are not restored (ported from ColorCharactersByWidth)
- Horizontal text is always turned vertical before processing
- **Conversion Check**: lists what will change, such as "3,000 → 三千" or "10 [TCY] ×3", and updates whenever an option or digit range changes
- A progress bar shows how far processing has gone
- Settings are saved and restored the next time
- Japanese / English UI switches automatically

### How to use

1. Select vertical text objects (or a range of text with the Type tool)
2. Run the script and choose the options in the dialog box
3. Check the result in "Conversion Check" and click OK

### Order of processing

1. Turn horizontal text vertical (always)
2. Paragraph settings (Solid mojikumi, Soft v2 kinsoku, justify with last line top)
3. Apply rotation to all of the text (rewritten characters inherit the rotated formatting)
4. Rewrite characters (full-width to half-width, separators, full-width digits, era years, kanji numerals)
5. Apply Tate-Chu-Yoko to each target

### Notes

- When the kanji and Tate-Chu-Yoko digit ranges overlap, kanji numerals win
- Without the era conversion, the four digits of a year such as 2026年 also get Tate-Chu-Yoko when [4 digits] is checked
- When a rewrite changes the number of characters (3,000 → 三千 etc.), the rewritten part takes the formatting of its first character
- Locked or hidden text is skipped

### Article

https://note.com/dtp_tranist/n/ndb55c5516fd6

### Update History

- v1.18.0 (2026-10-16) The [Mojikumi] list now includes the document's custom mojikumi sets (kept by name, so the same set is picked next time in documents that have it). Improved detection of word runs: digits with units (5mm and 10cm) stay eligible for [Units to squares]; with [Tate-Chu-Yoko for 2–3 capitals] on, all-capital runs such as NHK BS get Tate-Chu-Yoko word by word; words starting with digits such as 3D Printing are included; a capitalized word after a sentence-ending period starts a new run; wide spaces such as em spaces become half-width
- v1.17.2 (2026-10-16) Fixed runs of words (such as Design Week) being split word by word in Illustrator, which rotated "Week" or made it full-width. Words joined by no-break or thin spaces are now treated as one run too
- v1.17.1 (2026-10-16) Runs of words that include full-width alphanumerics, spaces or ＆ are now treated as one run too: not rotated, and made half-width (Design Ｗｅｅｋ → Design Week). Fixed a "there is no document" error after turning horizontal text vertical (frames recreated by the change are re-taken from the selection)
- v1.17.0 (2026-10-16) Added "Auto-size area type" to the Character and Paragraph panel. "Number hyphens" now uses ‐ (hyphen) instead of full-width －, which stayed horizontal in vertical text. Added "Leading" (auto-leading percentage, 175% by default) to the Character and Paragraph panel. The "magazines, ads" preset now uses TCY for numbers only (no capital abbreviations or !?). Removed "Include units"; numbers and units are no longer combined in Tate-Chu-Yoko. Leading is always set to Auto at the end (in Simple mode too). Runs of two or more words (Type & Layout) are no longer rotated, made full-width or changed
- v1.16.0 (2026-10-16) Added Check Coloring to Detailed mode (ported from ColorNumberByWidth)
- v1.15.0 (2026-10-16) Phone and postal numbers are no longer turned into ranges; added "Number hyphens". "Separators" is now a dropdown (、 / remove / keep). Added "Include units" for Tate-Chu-Yoko and the "Only 万・億・兆" kanji style. Years are always digit by digit. Rebuilt the built-in presets (seven)
- v1.14.1 (2026-10-15) Moved "Proportional metrics for full-width Latin" from Character Conversion to the Character and Paragraph panel (formerly Paragraph)
- v1.14.0 (2026-10-15) The "Short words to full-width" length is now linked to one below "Skip long Roman text" (5 and 4 by default)
- v1.13.1 (2026-10-15) Fixed shifted rewriting and Tate-Chu-Yoko positions in threaded text (the whole story is now one target). Every threaded frame is turned vertical
- v1.13.0 (2026-10-15) Added "Short words to full-width" (N letters or fewer) to Character Conversion in Detailed mode
- v1.12.0 (2026-10-15) Added "Skip long Roman text" (no rotation for runs of N characters or more). Simple mode now applies only the visible options
- v1.11.0 (2026-10-15) Added "Proportional metrics for full-width Latin" to Character Conversion in Detailed mode
- v1.10.0 (2026-10-15) Added Simple and Detailed modes (in Detailed mode, Character Conversion spans both columns with two columns inside) and "Units to squares". The check list is two rows taller. Removed "Convert horizontal text to vertical" and "Vertical text only" from the UI; horizontal text is now always turned vertical. Added the Paragraph panel (kinsoku and mojikumi) to Detailed mode, more tooltips, and cleaned up the code
- v1.9.0 (2026-10-15) Removed the Paragraph panel; Solid mojikumi, Soft v2 kinsoku and last-line-top justification are now always applied
- v1.8.0 (2026-10-15) Added "0 as 零" and "千 as 一千" to the kanji numerals
- v1.7.1 (2026-10-15) Centered the preset row in the dialog box
- v1.7.0 (2026-10-15) Added range marks to 〜, symbols to full-width, quotes to 〝〟, ! ? to full-width and single letters to full-width
- v1.6.0 (2026-10-15) Added presets (four built-ins, plus saving and deleting your own)
- v1.5.1 (2026-10-15) Made the dialog box more compact (shorter labels with tooltips, digit checkboxes and kanji styles on one row, a shorter check list)
- v1.5.0 (2026-10-15) Added "Convert horizontal text to vertical"
- v1.4.0 (2026-10-15) Added "Justify with last line top" to the Paragraph panel
- v1.3.0 (2026-10-15) Rotation now runs before the character rewriting (words such as "Web" or "PDF" were sometimes left unrotated). Added a progress bar
- v1.2.0 (2026-10-15) Added the Paragraph panel to apply the Solid (ベタ組み) mojikumi set and Soft v2 kinsoku
- v1.1.0 (2026-10-15) Tate-Chu-Yoko digits are now chosen with [2 digits] [3 digits] [4 digits] checkboxes. Added "Fixed-width forms by length"
- v1.0.0 (2026-10-15) Initial release
