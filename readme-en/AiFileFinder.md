# Find .ai files across several folders

[![Direct](https://img.shields.io/badge/Direct%20Link-AiFileFinder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/AiFileFinder.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiFileFinder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

You want to reopen an `.ai` file you made a while ago, but you cannot remember which year's folder it went into. Searching in the Finder takes you a long way from the canvas, and File > Open Recent only goes back a few dozen entries.

This script is a finder that filters `.ai` files across several registered folders by keyword and opens the selected file on the spot. The index is cached, so thousands of files do not keep you waiting.

### Features

- Searches across several registered folders, including everything in their subfolders
- Filters as you type (incremental search)
- Space-separated AND search, insensitive to case, width, hiragana/katakana, and voiced marks
- Your own keyword buttons: click to replace, Option-click to add to the current keyword
- Two lists: search folders on the left, file names with modified dates on the right
- Filter by year, or by a period (from / to)
- Sort by newest first or by name
- Exclusions: hide files whose name or folder contains a given word
- A file that is already open is brought to the front instead of being opened twice
- Option-double-click reveals in the Finder; double-clicking a folder opens that folder
- The index is cached for quicker later launches, and settings live in the Illustrator preferences

### How to use

1. On launch the index is built and the file list appears (only the first run takes a while)
2. Type in the keyword field to narrow the candidates
3. Pick a search folder on the left to list only that folder's files on the right
4. Select a file on the right and press `Enter`, double-click it, or click [Open]

Option-double-click reveals the file in the Finder instead of opening it. Double-clicking a folder on the left opens that search folder in the Finder.

Keyboard:

| Key | Action |
|---|---|
| ↓ | Move from the keyword field to the file list |
| Enter | Open the selected file (works from anywhere except the date fields) |
| Option + double-click | Reveal in the Finder instead of opening |
| Double-click a folder | Open that search folder in the Finder |
| Click the × in the keyword field | Clear the keyword and drop the filter |

### Options

**Keyword buttons**

Buttons shown in the filter panel. Enter one word per line in the preferences. The defaults are `icon` / `logo` / `ロゴ` / `アップデート`.

| Action | Result |
|---|---|
| Click | Replace the keyword with that word |
| Option + click | Add the word to the current keyword to narrow further |

The buttons wrap automatically to the width of the dialog.

**Year**

Filters by the year of the modified date. The list holds only the years that actually occur in the index, newest first. It is built when the dialog opens, so the choices do not shift while you type.

**Period**

Filters by a start and an end date. Any of these forms is accepted, and the field is rewritten as `2023/01/05` once committed.

```
2023.1.5    2023-01-05    2023/1/5    20230105
```

| Input | As a start date | As an end date |
|---|---|---|
| `2023.1.5` | From 2023-01-05 00:00 | Through 2023-01-05 23:59:59 |
| `2024.2` | From 2024-02-01 00:00 | Through 2024-02-29 23:59:59 |
| `2024` | From 2024-01-01 00:00 | Through 2024-12-31 23:59:59 |
| (empty) | No limit | No limit |

Anything that cannot be read is cleared on commit and the limit goes back off. `Enter` in a date field commits the value; it does not open a file.

**Exclusions**

Hides files whose name or folder contains one of these words. One word per line, and a single match is enough (OR). Exclusions apply immediately, without rebuilding the index.

The defaults are `note-cover-` / `backup` / `_old` / `_outlined`, meant for backup folders and for routine files that are numerous enough to get in the way.

**Preferences**

Opened with the [Preferences] button.

| Item | Content |
|---|---|
| Search Folders | The folders to search. Edit with [Add], [Remove], and [Reset] |
| Keyword Buttons | Words shown in the filter panel (one per line) |
| Exclusions | Words that hide a file (one per line) |
| [Rescan] | Apply the current edits, then rebuild the index |

Changing the search folders and clicking OK rebuilds the index for the new set. Their order is the order of the list on the left.

### The index cache

The result of scanning the search folders is stored in `~/Library/Application Support/AiFileFinder-index.txt`. A launch within 24 hours only reads that file, so the list appears at once.

It is rebuilt when:

- You click [Rescan] in the preferences
- You change the search folders
- More than 24 hours have passed since the last scan

### Notes

- Files added, moved, or deleted after the index was built are not reflected until it is rebuilt. Trying to open a file that is gone brings up a dialog saying so
- When many files match, only the first 300 are listed (rebuilding thousands of rows on every keystroke would make typing lag). The panel caption reads "showing N of M", so narrow it down with a keyword
- The list on the left holds search folders only. Subfolders are always searched, but they do not get their own rows
- Exclusions also match folder names. Adding `backup` hides everything under a folder with that name
- Files whose modified date cannot be read are left out of the year and period filters
- Option-double-click selects the file in the Finder when `/Applications/RevealInFinder.app`, built with Automator, is present. Without it, or outside macOS, it just opens the enclosing folder
- Search folders, keyword buttons, and exclusions are stored in the Illustrator preferences (`AiFileFinder.*`)
- The extension searched for is `.ai`. To change it, edit `FILE_EXT_RE` near the top of the script
- The default search folders are in `SEARCH_FOLDER_DEFAULTS` near the top of the script. Edit them for your own setup, or register your folders in the preferences

### Reference

Revealing a file in the Finder with Option-double-click uses the Automator app approach published in [Jibunyou Memo (@mute_racoon3631), "A real Reveal in Finder for Illustrator"](https://note.com/mute_racoon3631/n/n9e0e08f5d5f7).

ExtendScript has no way to select a file in the Finder, so the path is handed to the Automator app through a temporary file.

### Version history

- v1.0.0 (2026-08-27): First release
