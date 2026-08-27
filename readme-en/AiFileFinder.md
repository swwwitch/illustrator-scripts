# Find .ai/.svg files across several folders

[![Direct](https://img.shields.io/badge/Direct%20Link-AiFileFinder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/AiFileFinder.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiFileFinder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

You want to reopen an `.ai` file you made a while ago, but you cannot remember which year's folder it went into. Searching in the Finder takes you a long way from the canvas, and File > Open Recent only goes back a few dozen entries.

This script is a finder that filters `.ai` and `.svg` files across several registered folders by keyword and opens the selected file on the spot. The index is cached, so thousands of files do not keep you waiting.

### Features

- Searches across several registered folders, including everything in their subfolders
- Filters as you type (incremental search)
- Space-separated AND / OR search, insensitive to case, width, hiragana/katakana, and voiced marks
- `ai` / `svg` checkboxes pick which kinds are listed (`svg` starts unchecked)
- Your own keyword buttons: click to replace, Option-click to add to the current keyword
- Two lists: search folders on the left, file names with modified dates on the right
- Filter by year, or by a period (year and month)
- Sort by modified date or name, ascending or descending
- File names are listed without their extension
- Exclusions: hide files whose name or folder contains a given word; typing an excluded word as a keyword lifts that one exclusion
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

**AND / OR**

The radio buttons on the row under the keyword field decide how space-separated words are treated. AND is the default.

| Setting | Behaviour |
|---|---|
| AND | Keep files that contain **all** of the words; each word narrows the result |
| OR | Keep files that contain **any** of the words, which gathers spelling variants such as `ロゴ logo` |

With a single word the two behave the same.

**Extensions (ai / svg)**

The checkboxes at the right end of the AND / OR row pick which kinds of file are listed. Only `ai` starts checked.

Both extensions are in the index, so checking `svg` lists them at once, with no rescan. Unchecking both leaves the list empty.

To search for more extensions, add them to `FILE_EXTENSIONS` near the top of the script; the checkboxes follow that same order.

```js
var FILE_EXTENSIONS = [
    { ext: "ai",  isChecked: true },
    { ext: "svg", isChecked: false }
];
```

**Keyword buttons**

Buttons shown at the bottom of the filter panel. Enter one word per line in the preferences. The twelve defaults are `icon` / `logo` / `font` / `keyboard` / `アイコン` / `ロゴ` / `カラー` / `フォント` / `アップデート` / `ツール` / `パネル` / `線`.

| Action | Result |
|---|---|
| Click | Replace the keyword with that word |
| Option + click | Add the word to the current keyword to narrow further |

The buttons wrap automatically to the width of the dialog.

**Year and period**

Dropdowns on the row above the keyword buttons narrow the list by modified date.

```
Year: [All ▼]  Period: [None▼][01▼] – [None▼][12▼]
```

**Year** picks a single year. **Period** picks a start and an end as year and month: the start begins on day 1 at 00:00, and the end runs through the last day at 23:59:59 (leap years included).

The year dropdowns hold only the years that actually occur in the index, newest first. They are built when the dialog opens, so the choices do not shift while you type.

The month defaults to `01` on the left and `12` on the right, so picking a year alone covers it whole. While a year reads "None", the month next to it is dimmed.

| Setting | Range |
|---|---|
| 2024/01 – 2026/08 | From 2024-01-01 00:00 through 2026-08-31 23:59:59 |
| 2024/01 – None | From 2024-01-01 00:00, no upper limit |
| None – 2026/08 | No lower limit, through 2026-08-31 23:59:59 |

**Sort order**

Set on the row under the lists. `Modified` / `Name` radio buttons pick the key, and the ▲ (ascending) / ▼ (descending) buttons pick the direction.

```
Sort by: (•)Modified ( )Name  [▲][▼]
```

The default is Modified plus ▼, that is newest first. Files with the same date fall back to the name.

The ▲▼ buttons are drawn by the script, and the chosen one is filled.

Note that clicking a list header does not sort; a ScriptUI header cannot receive clicks.

**Exclusions**

Hides files whose name or folder contains one of these words. One word per line, and a single match is enough (OR). Exclusions apply immediately, without rebuilding the index.

The defaults are `note-cover-` / `backup` / `_old` / `_outlined` / `test` / `copy` / `名称未設定`, meant for backup folders and for routine files that are numerous enough to get in the way.

Typing an excluded word as a keyword lifts that one exclusion. With `backup` excluded, searching for `backup` still lists the backups. Only the word you typed is lifted, so `logo_old.ai` stays hidden when you search for `logo` while `_old` is excluded.

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
- Setting a start later than the end simply matches nothing; the two are not swapped for you
- Option-double-click selects the file in the Finder when `/Applications/RevealInFinder.app`, built with Automator, is present. Without it, or outside macOS, it just opens the enclosing folder
- Search folders, keyword buttons, and exclusions are stored in the Illustrator preferences (`AiFileFinder.*`)
- The extensions searched for are `.ai` and `.svg`. To change them, edit `FILE_EXTENSIONS` near the top of the script
- File names are listed without their extension, so an `.ai` and an `.svg` of the same name look alike. Set `SHOW_FILE_EXTENSION` near the top of the script to `true` to tell them apart
- The default search folders are in `SEARCH_FOLDER_DEFAULTS` near the top of the script. Edit them for your own setup, or register your folders in the preferences

### Reference

Revealing a file in the Finder with Option-double-click uses the Automator app approach published in [Jibunyou Memo (@mute_racoon3631), "A real Reveal in Finder for Illustrator"](https://note.com/mute_racoon3631/n/n9e0e08f5d5f7).

ExtendScript has no way to select a file in the Finder, so the path is handed to the Automator app through a temporary file.

### Version history

- v1.0.0 (2026-08-27): First release
