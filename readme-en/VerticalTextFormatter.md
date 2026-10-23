# Apply rotated Roman characters, Tate-Chu-Yoko and number rewriting to vertical text

[![Direct](https://img.shields.io/badge/Direct%20Link-VerticalTextFormatter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/VerticalTextFormatter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VerticalTextFormatter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Applies "Rotate Standard Roman Characters in Vertical Text" and Tate-Chu-Yoko to the selected vertical text in one go.
- It can also rewrite numbers as full-width digits, kanji numerals or Japanese era years, and turn separators into marks for vertical text.
- Choose the options in a dialog box and check the result in the [Conversion Check...] list before running.
- Rotation and Tate-Chu-Yoko cannot be set through the DOM, so a temporary (dynamic) action applies just those Character panel items. Font, size and other formatting stay as they are.

### Features

- Simple, Detailed and Wizard modes (switch at the top left of the dialog box; Simple by default)
  - **Simple**: shows only rotation, Tate-Chu-Yoko digits (2/3/4), fixed-width forms, separator conversion, and full-width single digits, non-TCY digits and letters
  - **Detailed**: shows every option
  - **Wizard**: choose what you want to do with checkboxes, and the matching detailed options are set for you (Detailed-mode settings are kept; unchecked options are off; presets are unavailable). Option-click a checkbox to turn all on (when all are on, only the clicked one stays on). OK is unavailable until at least one usable goal is checked
    - Rotate Roman text upright (runs of 5 or more characters stay sideways)
    - Stand short words and single letters upright (4 or fewer letters become full-width; takes priority over rotation)
    - Tate-Chu-Yoko for 2–3 digit numbers (fixed-width forms by length; unavailable with kanji numerals)
    - Stand other digits upright in full-width (unavailable with kanji numerals)
    - Convert numbers to kanji (style from the Detailed mode, including 零 and 一千)
    - Western years to Japanese eras
    - Tate-Chu-Yoko for abbreviations and !?
    - Turn units into squares
    - Use vertical-friendly marks (、 separators, ・ decimal points, 〜, full-width symbols, 〝〟, number hyphens)
    - Set up paragraphs for vertical text (Soft v2 kinsoku, half-width line-end punctuation, 175% leading, proportional metrics)
    - Fit area type frames to the text
    - Color characters afterwards to check
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
    - **Skip Roman text of [N] or more characters**: Roman text of at least the given number of characters (5 by default) stays sideways. Words joined by spaces or . ' & - count as one run ("Adobe Illustrator" is 17)
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
  - **Full-width for words of [N] or fewer letters** (Detailed mode): words of at most the given number of letters become full-width (OK → ＯＫ). The length is linked to one below "Skip Roman text of [N] or more characters" (5 and 4 by default). Unit squares and capital Tate-Chu-Yoko take priority
  - **Single letters to full-width**: a lone letter such as the A in Aランク
- Numbers
  - **Single digits to full-width**: keeps a single digit upright
  - **Non-TCY digits to full-width**: makes every digit not in Tate-Chu-Yoko full-width so it stands upright, including digits next to separators and in phone numbers, but not those made kanji (off by default; dims "Single digits to full-width" when on)
  - **Units to squares**: turns mm, cm, km, m, mg, g, kg, t, ml, dL, kL, L, cal, kcal, cc right after digits into ㍉ ㌢ ㌖ ㍍ ㍉㌘ ㌘ ㌕ ㌧ ㍉㍑ ㌥㍑ ㌔㍑ ㍑ ㌍ ㌔㌍ ㏄ (10cm → 10㌢), and m2, cm3, m² etc. into ㎡ ㎤. A space between the digits and the unit is removed (5 mm → 5㍉), and full-width units count too. Units followed by letters or / (km/h, mmHg) are left. Off by default because some fonts lack these glyphs
  - **Western years to Japanese era**: 2026年 → 令和8年. When the month and day follow, they decide which side of an era change the date falls on (2019年4月30日 → 平成31年, 2019年5月1日 → 令和元年); a bare year of an era change uses the new era. Year 1 is written 元年
  - **Kanji numerals to Arabic digits**: turns stray kanji numerals into digits before Tate-Chu-Yoko and full-width are decided (2019年四月30日 → 2019年4月30日, 三十日 → 30日). Covers runs of two or more kanji numerals and a single one before 月 日 時 分 秒; 一緒, 十分, 数十万, 何百 and the like stay. Unavailable while Convert to kanji numerals is on (to unify toward kanji, set its digit range to start at 1). Off by default
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
  - **Leading**: set as an auto-leading percentage (175% by default, 50–500%). In both modes, leading is always set to Auto at the end so it follows the type size, and auto kerning is always set to 0
  - Justification is always justified with the last line at the top
  - **Proportional metrics for full-width alphanumerics**: applies the OpenType Proportional Metrics feature to full-width letters and digits (after the rewriting, on by default)
  - **Height to a multiple of the type size**: sets the area-type frame height (the line length in vertical text) to the nearest whole multiple of the type size, inset excluded (on by default; always applied, in Simple mode too)
  - **Auto-size area type**: after processing, turns on Auto Size in Area Type Options so the frame fits the text (on by default and in every built-in preset; threaded text is skipped)
- Check Coloring (Detailed mode)
  - **Color characters after processing**: full-width alphanumerics and ！？ red, half-width blue, fixed-width glyphs (such as third- and quarter-width TCY) green, and the rest in the base color (black). Off by default. The colors are set in the CHECK_* variables at the top of the script. For checking only; the original colors are not restored (ported from ColorCharactersByWidth)
- Horizontal text is always turned vertical before processing
- **Conversion Check...**: lists what the current settings will change, such as "3,000 → 三千" or "10 [TCY] ×3", in a separate dialog box
- A progress bar shows how far processing has gone
- Settings are saved and restored the next time
- Japanese / English UI switches automatically

### How to use

1. Select vertical text objects (or a range of text with the Type tool)
2. Run the script and choose the options in the dialog box
3. Check the list of changes with [Conversion Check...] and click OK

### Order of processing

1. Return text that is already vertical to horizontal form (rotation and Tate-Chu-Yoko off; full-width alphanumerics, 、 separators, ・ decimal points, symbols, 〝〟, number hyphens and unit squares reverted; fixed-width glyphs reset — the same as HorizontalTextFormatter's defaults; range 〜 is kept, and kanji numerals and eras are reverted only in frames last processed with them on). Text that was already formatted can be run again with different settings
2. Turn horizontal text vertical (always)
3. Paragraph settings (Solid mojikumi, Soft v2 kinsoku, justify with last line top)
4. Apply rotation to all of the text (rewritten characters inherit the rotated formatting)
5. Rewrite characters (full-width to half-width, kanji numerals to digits, separators, full-width digits, era years, kanji numerals)
6. Apply Tate-Chu-Yoko to each target

### Notes

- For text that is already vertical, "Conversion Check" lists changes from the text after it is returned to horizontal form (a line appears even when the result is the same as before, such as ５ → 5 → ５)
- When the kanji and Tate-Chu-Yoko digit ranges overlap, kanji numerals win
- Without the era conversion, the four digits of a year such as 2026年 also get Tate-Chu-Yoko when [4 digits] is checked
- When a rewrite changes the number of characters (3,000 → 三千 etc.), the rewritten part takes the formatting of its first character
- Locked or hidden text is skipped

### Article

https://note.com/dtp_tranist/n/ndb55c5516fd6

### Update History

- v1.26.0 (2026-10-23) Added "Kanji numerals to Arabic digits" to Numbers in the Detailed mode: stray kanji numerals as in 2019年四月30日 become digits before Tate-Chu-Yoko and full-width are decided
- v1.25.3 (2026-10-16) Wizard fixes: loading a preset in the Detailed mode no longer resets the wizard checks, and 零 / 一千 now follow the Detailed-mode settings. Renamed "Turn Roman text sideways" to "Rotate Roman text upright". Turning on kanji numerals dims the Tate-Chu-Yoko and full-width digit goals, and OK is unavailable when no usable goal is checked
- v1.25.2 (2026-10-16) Option-clicking in the Wizard now turns all on, and a second option-click keeps only the clicked one on
- v1.25.1 (2026-10-16) Option-clicking a Wizard checkbox turns all of them on or off
- v1.25.0 (2026-10-16) Added the Wizard mode: choose what you want to do (12 goals such as turning Roman text sideways, Tate-Chu-Yoko for 2–3 digits, vertical-friendly marks) and the matching detailed options are set for you
- v1.24.5 (2026-10-16) Made the preset dropdown narrower (220 → 150 px)
- v1.24.4 (2026-10-16) Moved the Simple/Detailed switch to the top left and the presets to the top right. Moved [Conversion Check...] to the left of the button row
- v1.24.3 (2026-10-16) The kanji style radio buttons (二〇二六 / 二千二十六 / 二万五〇〇〇) are now stacked vertically
- v1.24.2 (2026-10-16) "Short words to full-width" now reads "Full-width for words of [N] or fewer letters" (clicking the text after the field also toggles it)
- v1.24.1 (2026-10-16) Swapped the columns in Detailed mode (left: Numbers and Character Conversion; right: rotation and Tate-Chu-Yoko, Character and Paragraph, Check Coloring)
- v1.24.0 (2026-10-16) "Conversion Check" is now a [Conversion Check...] button that opens the list in a separate dialog box. "Skip long Roman text" now reads "Skip Roman text of [N] or more characters"
- v1.23.0 (2026-10-16) When run again, kanji numerals and Japanese eras are also returned to Arabic numerals and Western years first (二〇二六 → 2026, 二千二十六 → 2026, 令和八年 → 2026年, 一月七日 → 1月7日). To keep phrases such as 数十万 and 四十七士 in the original text intact, processed text frames are tagged with the options, and only frames processed with kanji or eras on are reverted (frames processed with v1.22.0 or earlier have no tag and are not reverted)
- v1.22.0 (2026-10-16) Improved "Units to squares": a space between digits and the unit is removed (5 mm → 5㍉), full-width units (５ｍｍ) are included, dL, kL, cal, kcal and cc were added, and m2, cm3, m² become ㎡, ㎤. Units followed by letters or / (km/h, mmHg) are left as they are (60km/h used to become 60㌖／h). A slash between letters is no longer made full-width
- v1.21.1 (2026-10-16) "Separators" is a pop-up menu again
- v1.21.0 (2026-10-16) The script can now be run again on text that is already vertical. Vertical text is first returned to horizontal form, as with HorizontalTextFormatter's default conversions (rotation and Tate-Chu-Yoko off; full-width alphanumerics, 、 separators, ・ decimal points, symbols, 〝〟, number hyphens and unit squares reverted; fixed-width glyphs reset), and then formatted with the current settings. Era years, kanji numerals and range 〜 are not reverted
- v1.20.0 (2026-10-16) Auto kerning is now always set to 0 at the end (both modes; no option in the dialog box; manual kerning values are kept)
- v1.19.0 (2026-10-16) Added "Non-TCY digits to full-width" (also digits next to separators and in phone numbers; not those made kanji). "Separators" is now a vertical set of radio buttons. Check coloring no longer has swatches; the colors are set in the CHECK_* variables at the top of the script (the dialog box only turns it on or off)
- v1.18.1 (2026-10-16) Renamed "Proportional metrics for full-width Latin" to "Proportional metrics for full-width alphanumerics" and included full-width digits (０–９)
- v1.18.0 (2026-10-16) The [Mojikumi] list now includes the document's custom mojikumi sets (kept by name, so the same set is picked next time in documents that have it). Improved detection of word runs: digits with units (5mm and 10cm) stay eligible for [Units to squares]; with [Tate-Chu-Yoko for 2–3 capitals] on, all-capital runs such as NHK BS get Tate-Chu-Yoko word by word; words starting with digits such as 3D Printing are included; a capitalized word after a sentence-ending period starts a new run; wide spaces such as em spaces become half-width
- v1.17.2 (2026-10-16) Fixed runs of words (such as Design Week) being split word by word in Illustrator, which rotated "Week" or made it full-width. Words joined by no-break or thin spaces are now treated as one run too
- v1.17.1 (2026-10-16) Runs of words that include full-width alphanumerics, spaces or ＆ are now treated as one run too: not rotated, and made half-width (Design Ｗｅｅｋ → Design Week). Fixed a "there is no document" error after turning horizontal text vertical (frames recreated by the change are re-taken from the selection)
- v1.17.0 (2026-10-16) Added "Auto-size area type" to the Character and Paragraph panel. "Number hyphens" now uses ‐ (hyphen) instead of full-width －, which stayed horizontal in vertical text. Added "Leading" (auto-leading percentage, 175% by default) to the Character and Paragraph panel. The "magazines, ads" preset now uses TCY for numbers only (no capital abbreviations or !?). Removed "Include units"; numbers and units are no longer combined in Tate-Chu-Yoko. Leading is always set to Auto at the end (in Simple mode too). Runs of two or more words (Type & Layout) are no longer rotated, made full-width or changed
- v1.16.0 (2026-10-16) Added Check Coloring to Detailed mode (ported from ColorNumberByWidth)
- v1.15.0 (2026-10-16) Phone and postal numbers are no longer turned into ranges; added "Number hyphens". "Separators" is now a dropdown (、 / remove / keep). Added "Include units" for Tate-Chu-Yoko and the "Only 万・億・兆" kanji style. Years are always digit by digit. Rebuilt the built-in presets (seven)
- v1.14.1 (2026-10-15) Moved "Proportional metrics for full-width Latin" from Character Conversion to the Character and Paragraph panel (formerly Paragraph)
- v1.14.0 (2026-10-15) The "Short words to full-width" length is now linked to one below "Skip Roman text of [N] or more characters" (5 and 4 by default)
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
