# Move the paragraph at the cursor up

[![Direct](https://img.shields.io/badge/Direct%20Link-moveParagraphUp.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/moveParagraphUp.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/moveParagraphUp.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Swaps the paragraph containing the text cursor with the paragraph above it. It is Visual Studio Code's "Move Line Up" applied to paragraphs instead of display lines.
- A paragraph-based variant of moveLineUp.jsx by sky-chaser-high.
- The cursor stays inside the paragraph that moved, so running the script again keeps sending it up.

### Features

- Works with the caret placed anywhere in the paragraph; no need to select the whole paragraph
- Swaps the formatting too (font, size, paragraph style, and so on)
- Restores the cursor to the same position inside the paragraph, so the script can be repeated
- Does nothing on the first paragraph

### Usage

1. With the Type tool, place the cursor inside the paragraph to move (a selection inside the paragraph works too).
2. Run the script.
3. Run it repeatedly to move the paragraph further up.

Assigning the script to an action with a function key makes reordering paragraphs feel like working in a text editor.

### Notes

- The script runs only while text is being edited with the Type tool (the caret is inside the text). Selecting a text frame as an object does nothing.
- The swap goes through the clipboard, so running the script replaces the clipboard contents.
- When the destination is an empty paragraph (a lone return), the cursor position is not restored.
- `moveParagraphDown.jsx` moves a paragraph in the opposite direction.

### Update History

- v1.0.0 (2026-08-27) : Initial release
- v1.0.1 (2026-09-02) : Fixed "Error 21: undefined is not an object" raised while restoring the cursor after the swap (it read `contents` from the Story, which has no such property; the text now comes from the paragraph)
