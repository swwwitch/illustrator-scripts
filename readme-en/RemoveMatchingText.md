# Remove specified strings from text

[![Direct](https://img.shields.io/badge/Direct%20Link-RemoveMatchingText.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/RemoveMatchingText.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RemoveMatchingText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Removes the entered strings from text in one go.
- There are five fields, processed from top to bottom. Regular expressions are supported.
- Choose the scope: selected objects, the current artboard, or the entire document. Text inside symbols can be included.
- Characters are removed one by one, so the formatting of the remaining text (color, font, size, etc.) is kept.

### Features

- Up to five strings to remove (empty fields are ignored)
- Shows the number of matches in the scope for each field (updated as you type or change the scope)
- Regular expressions and case-insensitive search
- Scope: selected objects (including inside groups) / current artboard (text that overlaps it) / entire document
- Falls back to the entire document when nothing is selected
- Deletes text left empty by the removal
- Processes text in hidden or locked layers and objects (released only while processing, then restored)
- Processes text in symbols (by rewriting the symbol definition)
- Reports the removed count per field, and the numbers of changed text, deleted text and rewritten symbols
- Remembers the entries and settings for the next run

### How to use

1. Select the target, or run the script with nothing selected
2. Enter strings under "Text to Remove" (the match count appears to the right)
3. Choose the scope and options, then click OK

### Dialog

| Item | Description |
| --- | --- |
| Text to Remove | Strings to remove (five fields). Empty fields are ignored; fields are processed from top to bottom. The number on the right is the match count in the scope; "!" means an invalid regular expression |
| Regular expression | Treats the input as JavaScript regular expressions. `^` and `$` match the start and end of each paragraph |
| Ignore case | Treats upper- and lowercase letters as the same |
| Selected objects | Selected text (including inside groups). Unavailable when nothing is selected |
| Current artboard | Text that overlaps the active artboard, even partly |
| Entire document | All text in the document |
| Delete emptied text | Deletes text frames emptied by this run. Frames that were already empty and threaded text frames are kept |
| Search hidden layers | Includes text in hidden layers and objects. They are shown only while processing, then hidden again |
| Search locked layers | Includes text in locked layers and objects. They are unlocked only while processing, then locked again |
| Search symbols too | Includes text inside symbols in the scope |

### Notes

- The match count of each field ignores removals by the other fields. When the strings overlap, the actual number removed may differ.
- "Search symbols too" rewrites the symbol definition itself, so instances of the same symbol outside the scope change as well.
- Symbols are recreated and swapped, so symbol options such as the registration point and 9-slice scaling are reset.
- With "Search symbols too" on, symbols are expanded temporarily in the dialog to count matches. The dialog may open slowly in documents with many symbols.
- Clicking OK saves the entries and settings for the next run.

### Article

- [DTP Transit (Japanese)](https://note.com/dtp_tranist/n/nec5dfffce709)

### Changelog

- v1.0.0 (20260926) : Initial release
