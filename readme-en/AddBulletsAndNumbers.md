# Add Bullets and Numbers to Line Heads

[![Direct](https://img.shields.io/badge/Direct%20Link-AddBulletsAndNumbers.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AddBulletsAndNumbers.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBulletsAndNumbers.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that adds a bullet symbol or sequential numbers to the head of each line in the selected text frames.
- When the dialog opens, the type, glyph, number style and delimiter are detected from the existing leading markers, so you can carry on from what is already there.
- Every setting is reflected in a live preview, and Cancel restores the original state.
- Marker advance is driven by tab stops, so wrapped lines can be aligned with the body text.

### Main Features

- List Type
  - **Bullets** / **Numbered List** / **None** ("None" removes the leading markers and the tab stops)
- Bullet Symbol
  - Choose from • - ● ○ ◎ ■ □ ◆ ◇
  - Each glyph brings its own default scale and body position (•/- place the body at 0.8x the font size)
- Number Style
  - **Numbers** / **Circled (white)** (1-20) / **Circled (black)** (1-20) / **ABC** / **abc**
  - Circled numbers fall back to plain digits beyond 20, with a warning before you run
  - "Zero padding" pads with leading zeros to the largest width (numbers style only)
  - "Start No." sets the first number (default: 1)
  - "Restart each frame" restarts numbering at the head of every selected frame
- Delimiter
  - None / . / ： / | (switches to "None" automatically when a circled style is chosen)
- Position (Tab Stops)
  - Adjust **Marker** (where the number sits) and **Body** (where the text starts), starting from font-size based defaults
  - Defaults are derived per type: numbers = 1.5/2.0x, ABC/abc = 1.0/2.0x, bullets and circled = 1.2x
  - Alignment of the marker column: left / center / right (numbers, ABC and abc only; numbers default to right, ABC/abc to center)
  - Bullets and circled numbers use a single tab stop (the body position)
- Marker Format
  - Font family and style ("Japanese fonts only" narrows the list)
  - Scale (applied equally to horizontal and vertical) and baseline shift
  - Separate colors for the marker/number and the delimiter (click a swatch to open the color picker)
- Paragraph Settings
  - Leading and space after (left blank = leave unchanged)
  - Leading is not written as a fixed value: the ratio to the font size is derived per paragraph and stored as the auto-leading amount (%), so the leading follows along if the font size changes later
  - "Justification" is chosen with icon buttons (left / center / right / justify with the last line left / justify all)
  - The default follows the text kind: left for point text, justified with the last line left for area text
  - "Align wrapped lines" aligns the second and later lines with the body position
  - The indent value follows the body position automatically
- Reorder Lines
  - Select several rows in the list and move them to the top, up, down, or to the bottom
  - Sort all rows by name, by number, or by number descending (rows without a number go last)
- "Show Hidden Characters" toggles display of tabs and line breaks
- "Reset" clears scale, baseline shift, color, indent and tab stops at once
- Numeric fields step with the arrow keys (Shift = ±10, Option = ±0.1)
- Automatic Japanese / English switching

### Workflow

1. Collect text frames from the selection (recursing into groups) and order them top to bottom (then left to right)
2. Snapshot the text, character attributes, paragraph settings and tab stops
3. Show the dialog and initialize it from the markers already present
4. On every change, roll back the previous preview, write the snapshot back, and reapply with the current settings
5. OK commits (the preview is rolled back first and the settings are applied once, so a single undo step is added)
6. Cancel, or closing the dialog, restores the original state from the snapshot

### Not Supported

- No document open, nothing selected, or no text frame in the selection (a warning is shown and the script exits)
- Empty (whitespace-only) lines: they get neither a marker nor a number
- Line reordering: after committing it can only be undone through Illustrator's own undo; neither "Reset" nor Cancel restores the original order
- Character attributes such as composite fonts, OpenType features and character rotation: they are not part of what is restored when the body text is rebuilt

### Notes

Leading middle dots (・ ･ ·) and a hyphen (-) followed by a tab or space are stripped on run.
Hand-typed heads such as "1. " or "A) " are stripped too when a delimiter is present.
When a digit follows the delimiter, as in "12.5", the text is treated as body text and left intact.

### Update History

- v1.2.0 (2026-08-18): Added justification buttons to Paragraph Settings (the default follows the text kind: left for point text, justify-last-line-left for area text). Leading is now applied as an auto-leading amount (%) instead of a fixed value. Fixed tab stops, indents and space-after not being restored when Cancel followed "Reset". Right-aligned the labels in Marker Format and Paragraph Settings. Added tooltips to "None", the bullet symbol panel and the circled number styles, and expanded the ones on "Start No." and "Reset". Dimming the Body position now dims its label and unit as well. Sped things up by no longer rebuilding the line list on every preview and by reusing the Japanese-font keyword table. Unified the dialog margins and spacing. Preview updates are now deferred while typing or stepping through tab stop / scale values, and applied once the input settles
- v1.1.1 (2026-06-10)
- v1.0.0 (2026-05-30)
