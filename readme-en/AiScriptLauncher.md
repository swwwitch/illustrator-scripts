# Filter scripts by keyword and run them

[![Direct](https://img.shields.io/badge/Direct%20Link-AiScriptLauncher.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiScriptLauncher.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiScriptLauncher.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Once your script collection grows, just finding the one you want takes time. This launcher filters the `.jsx` / `.js` / `.jsxbin` files in a chosen folder by keyword and runs the selected script on the spot.

Folders and file names are shown in two side-by-side lists, so you reach the script you want without walking the hierarchy.

### Features

- Incremental search that narrows the candidates as you type
- Two lists: folders on the left, file names on the right. Picking a folder fills the right list with its contents
- AND search with space-separated terms
- A clear (×) button next to the keyword field empties it in one click, and is dimmed while the field is empty
- Words that appear often in the filtered results become one-click buttons, so you can drill down step by step
- Option double-click reveals a script in the Finder; double-clicking a folder opens it
- Nested subfolders (such as `artboard/backup`) can be excluded in one click
- The target folder and keyword settings are remembered in Illustrator's preferences, so they survive a restart

### Usage

1. Choose your script folder on the first run. Later runs reopen the remembered folder
2. Type in the keyword field to narrow the candidates
3. Pick a folder in the left list to list only that folder's file names on the right
4. Select in the right list and press `Enter`, double-click, or click **Run**

Option double-click reveals the script in the Finder instead of running it. Double-clicking a folder in the left list opens that folder in the Finder.

Keyboard:

| Key | Action |
|---|---|
| Down | Move from the keyword field to the file name list |
| Enter | Run the selected script (works wherever the focus is) |
| Option + double-click | Reveal in the Finder instead of running |
| Double-click a folder | Open that folder in the Finder |
| Click the × in the keyword field | Clear the keyword and drop the filter |

### Options

**Keyword buttons**

Script names are split into words, and the most frequent ones are laid out as buttons. `SmartDistributor.jsx` yields `Smart` and `Distributor`. Camel case, hyphens, underscores and digits are all treated as separators.

The buttons are **recomputed on every filter change**. Words the query already covers are dropped, since pressing them would narrow nothing, so what remains is the set of useful next steps.

```
(no keyword)   469 files  Text  Smart  Align  Artboard  Group …
  -> click Text
text           126 files  Auto  Type  Area  Leading  Kerning …
  -> option-click Align
text align       2 files  (no buttons left)
```

| Action | Result |
|---|---|
| Click | Replace the keyword with that word |
| Option + click | Append the word to the current keyword to narrow further |

**Preferences** lets you change how the buttons are chosen.

| Field | Meaning | Default |
|---|---|---|
| Occurrences | Minimum number of files a word must appear in | 4 |
| Keywords | Maximum number of buttons | 10 |

The number fields step by ±1 with the arrow keys, or ±10 with shift. Settings are stored in Illustrator's preferences and survive a restart.

**Include subdirectories**

When off, folders more than one level below the target folder are dropped from the list. This is aimed at hiding nested backup folders.

```
root/
├ Foo.jsx              shown
├ artboard/            shown
│  ├ Bar.jsx           shown
│  └ backup/           excluded
│     └ Bar-v2.jsx     hidden
```

**Preferences dialog**

Opened with the **Preferences** button. It holds the target folder and the keyword button rules together.

| Item | Meaning |
|---|---|
| Target folder | Shows the current path. **Change Folder** picks a new one |
| Full path | When off, the home folder is abbreviated to `~` |
| Occurrences / Keywords | The keyword button rules described above |

Changing the folder and pressing OK rebuilds the list against the new folder.

### Notes

- The launcher window itself does not show the target folder. Open **Preferences** to see where you are pointed
- Option double-click reveals the file with the Automator app at `/Applications/RevealInFinder.app` when it is installed. Without it, or outside macOS, it just opens the enclosing folder
- If this launcher itself lives inside the target folder, it is left out of the list
- File names written only in Japanese produce no keyword buttons, since the extractor looks for ASCII words
- The target folder and keyword settings live in Illustrator's own preferences (`AiScriptLauncher.*`). No external settings file is created
- The keyword, the “Include subdirectories” and “Full path” states and the list selections carry over only until Illustrator quits. They are not written to the preferences, so a restart brings back the defaults
- On launch the script turns `ShowExternalJSXWarning` off, so running scripts from outside the Scripts folder does not raise a warning every time. This preference applies to Illustrator as a whole

### Article

- [Introducing ScriptLauncher (note, in Japanese)](https://note.com/dtp_tranist/n/n86fe7e6251ec)

### Related articles

- [【Illustrator】スクリプトファイルの格納場所と実行方法](https://note.com/dtp_tranist/n/n9de5e22a4854) — where to keep script files, and five ways to run them
- [【Illustrator】面倒くさがり屋さんのためのスクリプトの管理と実行](https://note.com/dtp_tranist/n/n5da05c2e8c4e) — managing scripts on cloud storage and running them with Keyboard Maestro

Both cover where scripts should live and how to run them without a launcher.

### Reference

Revealing a file with it selected in the Finder (option double-click) uses the Automator bridge published in [「真の「Finderで表示」をイラレでも」 by 自分用メモ (@mute_racoon3631)](https://note.com/mute_racoon3631/n/n9e0e08f5d5f7).

ExtendScript has no way to reveal a file with selection, so the path is handed to an Automator app through a temporary file.

### Update History

- v1.4.2 (2026-08-30): The keyword and the list selections now carry over between runs within an Illustrator session
- v1.4.1 (2026-08-27): Added a clear (×) button to the keyword field
- v1.4.0 (2026-08-26): Switched to two side-by-side lists; added keyword buttons, Finder reveal and the Preferences dialog
- v1.0.0 (2025-11-13): Initial release
