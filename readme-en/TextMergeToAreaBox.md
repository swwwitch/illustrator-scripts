# TextMergeToAreaBox

[![Direct](https://img.shields.io/badge/Direct%20Link-TextMergeToAreaBox.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextMergeToAreaBox.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextMergeToAreaBox.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Rebuilds text items that were split apart — as happens when a PDF is opened in Illustrator — into a single area text frame.
- Items with similar Y positions are treated as one line, and each line is merged from left to right.
- Font, size, width and leading are inherited from the original text.
- The original text items are deleted and replaced by the generated area text.

### How to use

1. Select every text item you want to rebuild.
2. Run the script.
3. The selection is replaced by the area text, which is left selected.

Non-text objects in the selection are ignored. If the selection contains no text at all, the script shows "No convertible text found." and stops.

### Merging and line break rules

When lines are joined, the end of the previous line and the start of the next one decide what happens.

| End of previous line | Start of next line | Result |
| --- | --- | --- |
| Alphanumeric or `)` | Alphanumeric or `(` | A space is inserted |
| Alphanumeric plus a hyphen | Alphanumeric or `(` | The hyphen is removed and the words are joined |
| "。", "！" or "？" | — | A line break is inserted |
| ".", "!" or "?" | — | A line break is inserted |
| Anything else | — | Joined without a line break |

Sentence-ending punctuation triggers a break in both Japanese and Western text. As a result, a line ending with an abbreviation such as "Fig." or "etc." also gets a break even though the sentence continues.

### Inherited formatting

| Item | Description |
| --- | --- |
| Font and size | Inherited from the topmost line |
| Width | The width of the combined bounding box, reduced by one character (the font size); at least one character is kept for narrow selections |
| Height | The height of the combined bounding box, extended downward when the last line would overflow (the top edge stays put) |
| Leading | The Y gap between the first and second line (auto leading is turned off) |
| Justification | Justify with the last line left-aligned |
| Kinsoku | Set to "Soft_v2" (applied to the whole text, so it also shows in the Paragraph panel). Falls back to "Soft" on versions without it |
| Orientation | Vertical text is unified to horizontal |

If the measured leading is smaller than the font size, 1.2 times the font size is used instead.

### When only one line is selected

With a single line, the script does not create area text: it merges the items from left to right into one point text and left-aligns it.

### Settings

These can be adjusted in the "User settings" block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `LINE_Y_THRESHOLD` | 5 | Y threshold for grouping items into the same line (pt) |
| `MIN_LEADING_RATIO` | 1.2 | Minimum leading ratio against the font size |
| `DEFAULT_KINSOKU` | `"Soft_v2"` | Kinsoku to apply (`None` / `Hard` / `Soft` / `Soft_v2`) |
| `FALLBACK_KINSOKU` | `"Soft"` | Kinsoku used when `DEFAULT_KINSOKU` is unavailable |
| `AREA_TEXT_JUSTIFICATION` | Justify, last line left | Justification to apply |
| `MAX_HEIGHT_GROW_STEPS` | 20 | Maximum number of times the frame is extended downward (one leading per step) |

Adjust `LINE_Y_THRESHOLD` if lines are split unexpectedly, or if separate lines end up merged into one.

### Process flow

1. Verify that a document is open
2. Collect only the text frames from the selection and sort them from top to bottom
3. Group items with similar Y positions into the same line
4. Merge each line from left to right into a single frame
5. Join all lines into one string following the line end / line start rules
6. Build a rectangle from the combined bounding box, convert it to area text and pour the string in
7. Apply font, size, leading, justification and kinsoku
8. Extend the frame downward if the text overflows, then delete the source frames

### Notes

- The original text items are deleted. Duplicate them first if you need to keep them.
- Leading is measured from the apparent Y gap, so space before / after paragraphs is not supported.
- Vertical text is converted to horizontal.
- Lines are detected from the frame position (top left), so items on the same line may be split apart when their font sizes differ greatly.
- A line ending with an abbreviation ("Fig.", "etc." and so on) gets a break even though the sentence continues.
- Because the frame is extended to clear overflow, the bottom edge of the generated area text can end up below the original bounding box.

### Related script

- [TextMergeToAreaBox-tab.jsx](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextMergeToAreaBox-tab.md) : Joins the items within a line with tabs instead. Useful for rebuilding tables.

### Changelog

- v1.0 (2025-07-18) : Initial release
- v1.1 (2025-07-19) : Added support for single line text, set kinsoku rules
- v1.2 (2025-07-20) : Added handling for line breaks after English words
- v1.2.1 (2026-06-18) : Refactored (IIFE wrap, function split, renaming); single line now forced horizontal; bounding box no longer depends on selection state; added guard for no open document
- v1.3.0 (2026-08-13) : Changed kinsoku to "Soft_v2" (applied to the whole text so it shows in the Paragraph panel, falling back to "Soft" where unavailable); lines made up of single-byte characters now break at sentence ends; the frame is extended downward when the last line overflows; added a minimum width so narrow selections no longer fail

### Script info

- Version: v1.3.0
- First release: 2025-07-18
- Last updated: 2026-08-13
- Article: https://note.com/dtp_tranist/n/ne8d31278c266
