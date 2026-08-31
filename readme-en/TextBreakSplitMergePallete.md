# Text Processing

[![Direct](https://img.shields.io/badge/Direct%20Link-TextBreakSplitMergePallete.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextBreakSplitMergePallete.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextBreakSplitMergePallete.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Text problems in Illustrator tend to be small and tedious. Text pasted from a PDF ends up scattered across separate frames, paragraph breaks and forced breaks are mixed together, stray spaces sit between CJK and Latin characters, lines need reordering one by one.

Each of those is a ten-line script on its own — but once you have a dozen of them, *finding the right script* becomes the tedious part, and you run out of keyboard shortcuts. So all of it lives in a single palette here. The palette stays open, so you can change the selection and keep pressing buttons.

## How to use

1. Select a text frame (or a group containing text).
2. Run the script.

Every button runs immediately. There is no preview and no Apply button — use `Cmd + Z` to undo. Press `Esc` to close the palette.

If the palette is already open, the script brings it to the front instead of launching a second copy.

### Status

The top of the palette always shows what is currently selected.

| Field | Meaning |
| --- | --- |
| Target Texts | Total number of text frames being targeted |
| Point Type | How many of them are point type |
| Area Text | How many of them are area text |
| Paragraph Breaks | Number of paragraph breaks |
| Forced Breaks | Number of forced breaks |
| Tabs | Number of tabs |

**Buttons for operations that cannot run are dimmed automatically.** With no forced breaks, *Forced Breaks to Paragraph Breaks* is disabled; with a single text frame, the merge buttons are disabled. The enabled state tells you what is available.

## Basic tab

### Breaks

*Merge All into One Line* is the greedy one: it merges the frames vertically, drops empty lines, removes every paragraph and forced break, then trims leading/trailing spaces, CJK–Latin spaces and repeated spaces.

When a break sits between two Latin words, it becomes a space so the words do not run together (`fast` and `car` give `fast car`, not `fastcar`). Breaks between CJK characters, or between CJK and Latin, are still closed up.

| Button | Action |
| --- | --- |
| Merge All into One Line | Merge, drop empty lines, remove all breaks, tidy spaces |
| Line Breaks Only | Remove paragraph breaks (forced breaks too when *Include Forced Breaks* is on) |
| Insert Line Break After Each Character | Insert a break after every character |
| At Specified Characters | Break right after each character listed in the text field |
| At Character Count | Break every N characters |
| Forced Breaks to Paragraph Breaks | Convert forced breaks into paragraph breaks |
| Paragraph Breaks to Forced Breaks | Convert paragraph breaks into forced breaks |

*At Specified Characters* starts with `、。，．｡､,.!?！？`, and the field is freely editable. No break is added where one already follows, and none is added when nothing but whitespace remains after that character (i.e. a sentence-final period).

*At Character Count* defaults to 35. Turn on the *Forced Break* checkbox next to it to wrap with forced breaks instead of paragraph breaks. Character counting restarts on each existing line.

### Split

| Button | Action |
| --- | --- |
| Split by Line Breaks | Split into a separate text frame per line |
| Split by Line Breaks (Keep Style) | Same, keeping character formatting and position |
| Split by Tabs | Split each paragraph at its tab positions |
| Keep Style (split by character) | One frame per character, formatting kept |
| Ignore Style (split by character) | One frame per character, formatting reset |

The two line-break splits work differently under the hood. The first duplicates the frame, replaces its contents and positions the copies from the leading value. The second duplicates each paragraph with `TextRange.duplicate` and re-aligns it against the bottom edge (left edge for vertical text) recorded before the split. Use the second one when formatting matters.

*Split by Tabs* measures **how far the character after each tab sits from the start of the paragraph** rather than reading the tab itself. The offsets come from outlining a duplicate, so table-like text keeps its visual alignment (when the measurement is not possible it falls back to spacing each piece half an em after the previous one).

Character-level splitting duplicates the frame, converts it to outlines, reads the bounding box of each glyph, and fits a new frame to it. If outlining fails it falls back to accumulating character widths. Breaks, tabs and spaces (half- and full-width) never become frames.

Split results are left as separate frames. **Hold Option (Alt) while clicking to collect them into a single group** — this applies to all five split buttons.

*Ignore Style* is not a full reset: **the font and size of the first character are kept.** Fill becomes black, baseline shift / rotation / tracking go to 0, and horizontal/vertical scale return to 100%.

### Concatenate

| Button | Action |
| --- | --- |
| Vertical | Merge top to bottom into a single text |
| Merge Horizontally (Keep Rows) | Merge each row left to right, keeping the rows as separate frames |
| Merge Horizontally (Merge Rows) | Merge horizontally, then combine the rows into one text |
| Format PDF Text | Merge horizontally and output as area text |

*Vertical* breaks each frame into lines, rebuilds provisional Y positions from the original leading, and re-sorts. Frames are joined in **visual order**, not selection order.

*Merge Horizontally (Merge Rows)* and *Format PDF Text* target text poured in from a PDF, and join rows with these rules:

- Forced breaks are removed first
- If a line ends with an alphanumeric plus a hyphen and the next line starts with an alphanumeric, **the hyphen is dropped** and the lines are joined
- If a line ends and the next begins with alphanumerics, **a word space is inserted**
- A break is kept **only** when the line ends with `。！？` (or with `.!?` when the text is not pure ASCII)

In short, lines that are merely wrapped get joined; lines where a sentence ends keep their break. Leading is restored from the Y difference between the merged rows, and kinsoku is set to the loose setting.

*Merge Horizontally (Merge Rows)* picks point type or area text from the selection (area text wins when the two are mixed). *Format PDF Text* always outputs area text.

Merge results are left ungrouped as well. These are visual approximations: per-line formatting differences and rotation are not preserved. Two frames count as the same row when their Y coordinates differ by 5 or less.

## Cleanup tab

### Tabs and spaces

| Button | Action |
| --- | --- |
| Remove Tabs | Delete tabs |
| Tabs to Spaces | Convert tabs to half-width spaces |
| Leading/Trailing Spaces | Remove spaces at the start and end of each line |
| Remove Spaces Between CJK and Latin | Remove spaces between CJK and Latin text |
| Collapse Spaces | Collapse runs of spaces into one |
| All at Once | Run trim, CJK/Latin and collapse in order |
| Remove All | Remove every space, half- and full-width |
| Space After . and , | Insert a space right after a period or comma |

*Remove Spaces Between CJK and Latin* precisely means "**keep only the spaces that sit between two alphanumerics, and delete the rest**". Word spacing inside Latin text survives; spaces inside Japanese text do not.

*Collapse Spaces* handles half-width and full-width runs separately: a run of half-width spaces becomes one half-width space, a run of full-width spaces becomes one full-width space.

*Space After . and ,* skips cases where the next character is whitespace, a digit, a period or a comma, so `3.14` and `1,000` stay intact.

### Spaces & symbols

Pick a *Before* and an *After* option with the radio buttons, then press *Convert*.

| Option | Target |
| --- | --- |
| Space | Half-width and full-width spaces |
| Underscore | `_` |
| Hyphen | `-` |

*Convert* is dimmed while Before and After are the same.

### Convert and remove list markers

- *Fullwidth to Halfwidth*: full-width digits and Latin letters become half-width
- *Halfwidth Kana to Fullwidth*: half-width kana become full-width (including dakuten/handakuten composition and `｡｢｣､･`)
- *Bullet List*: removes leading markers (`・` `･` `·` `•` `◦` `●` `○` `◎` `□` `■` `◆` `◇` `✓` `-` `*`)
- *Number List*: removes leading numbering (`1.` `１．` `①` `❶` `a.` `一.`; delimiters `.` `．` `:` `：` `|`)

Both also handle the "tab + marker + tab" form produced by [AddBulletsAndNumbers.jsx](AddBulletsAndNumbers.md), and leading indentation is removed along with the marker.

A few exceptions keep body text safe: `-5℃` and `*important` are left alone because no space follows the symbol, and `12.5` is not treated as numbering because a digit follows the period. Lines without a delimiter (`Apple is red`) are left alone too — but note that `Mr. Smith` does become `Smith`, so take care with English text.

**Only markers that exist as characters are affected.** Markers drawn from paragraph attributes are not part of the text and cannot be removed here. To tell them apart, look at the list in the *Line Edit* tab: **if the marker shows up there, it is a character and can be removed.**

## Line Edit tab

The list box on the left shows each line of the selected text. **It only works while exactly one text frame is selected.**

- *Up* / *Down*: swap the selected line with its neighbour
- *Add*: append a line entered in a dialog
- *Edit*: edit the selected line (double-clicking the list works too)
- *Delete*: delete the selected line (with confirmation)

Every list operation is written straight back to the text frame.

The buttons on the right apply to all lines at once.

| Button | Action |
| --- | --- |
| Sort | Sort lines by character code |
| Sort (Length) | Sort lines from shortest to longest |
| Reverse Order | Reverse the line order |
| Remove Duplicates | Remove duplicate lines, keeping the first occurrence |
| Remove Empty Lines | Remove empty lines |

## Convert tab

**Every button shows a preview of its result on the right**, so you can see the outcome before pressing it. The preview uses the first selected text and is truncated past 40 characters.

### Letter case

Letter case conversion.

| Button | Action |
| --- | --- |
| UPPERCASE | UPPERCASE |
| lowercase | lowercase |
| Capitalize Words | Capitalize Words |
| Sentence case | Sentence case |
| Title Case | Title Case |

*Title Case* is a Title Caps implementation: articles, prepositions and conjunctions (a, the, of, and, in …) stay lowercase, except at the start and end. Exceptions such as `v.` / `vs.`, `AT&T` and `Q&A` are handled.

*Capitalize Words* and *Sentence case* lowercase the whole text first, so all-caps input produces the intended result.

### Kana conversion

Converts freely between hiragana, katakana and halfwidth kana. The source does not matter: *Hiragana* accepts both fullwidth katakana and halfwidth kana, and *Katakana* accepts both hiragana and halfwidth kana.

| Button | Action |
| --- | --- |
| Hiragana | Katakana and halfwidth kana to hiragana |
| Katakana | Hiragana and halfwidth kana to katakana |
| Halfwidth Kana | Hiragana and katakana to halfwidth kana |

Anything involving halfwidth kana is first normalised to fullwidth kana and then converted. Voiced and semi-voiced marks are composed into a single character in fullwidth (`ｶﾞ` → `ガ`, `ｳﾞ` → `ヴ`) and decomposed into two in halfwidth (`パ` → `ﾊﾟ`). `｡｢｣､･` and the long vowel mark are converted as well.

The long vowel mark `ー` is left as is. `ヵ` and `ヶ` are not converted to hiragana, so counters such as `ヶ月` stay intact.

### Alphanumeric

Converts digits between fullwidth, halfwidth and kanji numerals. **All five accept halfwidth digits, fullwidth digits or kanji numerals as input.**

| Button | Action |
| --- | --- |
| Halfwidth Digits | Normalise to halfwidth digits |
| Fullwidth Digits | Normalise to fullwidth digits |
| Fullwidth if Single Digit | Single digits to fullwidth, runs of two or more to halfwidth |
| Kanji Numerals | To kanji (`2026` → `二〇二六`) |
| Kanji (Positional) | To positional kanji (`1234` → `千二百三十四`) |

*Fullwidth if Single Digit* is meant for vertical text: each run of digits is counted as a unit, so `8月` becomes `８月` while `31日` stays `31日`.

*Kanji Numerals* replaces digit by digit without positional markers (十/百/千); `0` becomes `〇`. *Kanji (Positional)* writes them with positional markers up to 兆 (`25000` → `二万五千`). Both accept fullwidth digits.

Kanji numerals are read back into arabic numerals first, then converted. Both digit-by-digit (`二〇二六`) and positional (`千二百三十四`, up to 億) notation is understood, so `三十一日` → *Kanji Numerals* → `三一日` and `二〇二六年` → *Kanji (Positional)* → `二千二十六年` work as well.

**It works purely on the characters, so words that merely contain a numeral are converted too: `十分` becomes `10分` and `一部` becomes `1部`.** Narrow the selection before running it.

## Hidden characters button

*Show/Hide Hidden Characters* at the bottom left of the palette toggles Illustrator's "Show Hidden Characters". While it is on, the button label turns blue. Use it to see at a glance whether a break is a paragraph break or a forced break.

**If you turned it on with this button, it is turned back off automatically when the palette closes.**

## Targets

- Text frames
- Groups containing text frames (processed recursively)
- Text selections that Illustrator returns as a TextRange

## Notes

An Illustrator palette loses its connection to the DOM while it is displayed, so every actual text operation is delegated to the main engine over BridgeTalk. The delegated code is wrapped with `encodeURIComponent` before sending and restored on the other side, because BridgeTalk escapes backslashes in the message body and would otherwise break `\r`.

Status updates have their own constraint: Illustrator 30.x has no timer API (`app.scheduleTask` / `setTimeout`), so the selection cannot be polled continuously. Instead the selection is re-read at the moment the user comes back to the palette — when it regains focus and on mouse-over. A lightweight payload carrying only the counting functions is used for that, rather than sending everything each time.

Split and merge results are left ungrouped (Option-click a split button to group them). Groups left holding only one item are ungrouped automatically when the palette closes.

### note

- [【Illustrator】テキストの改行削除、改行で分割、連結をまとめてカバーするスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/nf6f34559ba46)
