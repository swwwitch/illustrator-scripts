# Look up code for menu commands, tools and preferences

[![Direct](https://img.shields.io/badge/Direct%20Link-AiCommandPrefLookup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiCommandPrefLookup.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCommandPrefLookup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Pick items from the list built into the script and get `app.executeMenuCommand()`, `app.selectTool()` or `app.preferences` get/set code. Use it as a reference for looking up IDs and keys from menu or setting names.

### Features

- Kind: switch between Menu commands, Preferences and Tools
  - Menu commands: `app.executeMenuCommand('…');`
  - Tools: `app.selectTool('…');`
  - Preferences: two lines, `app.preferences.get…Preference('…');` and `set…Preference('…', value);` (Boolean / Integer / Real / String by type; the sample value comes from the actual preference file)
- Language: show menu and setting names in Japanese or English
- Category: filter by top-level menu (File, Edit, Object…) or preference section
- Keyword: filters names and IDs/keys as you type (case-insensitive, regular expressions allowed)
- With several items selected, the code is listed in list order
- Add menu names as comments: appends the menu name, such as `// File > New...`, to each line (off by default)
- Memo: shows what a preference value means (such as the unit codes of `rulerType`) and caveats such as "takes effect after a restart"
- Recheck…: finds items that need updating after an Illustrator upgrade (see below)

### How to use

1. Run the script (no document needs to be open)
2. Choose the kind and language, then filter by category or keyword
3. Select an item to see its name, memo and code
4. Select all in the code field, copy, and paste into your script

### Recheck

Recheck… reads the settings folder of the running Illustrator and shows how it differs from the list.

- Compares with the keyboard shortcut files (.kys): lists menu commands and tools that may have been renamed or removed, and IDs missing from the list
- Compares with the preference files (Adobe Illustrator Cloud Prefs / Adobe Illustrator Prefs): lists keys missing from the list and type mismatches
- Opens the reference URLs in a browser (the Adobe Community thread, Ai Command Palette, sttk3's Notion database, Ten A's list)
- Opens the settings folder and saves the result as a text file

When the Illustrator version the list was checked with differs from the running one, a notice appears at the top of the result.

### Notes

- The list is built into the script. Its source is a text file combining the menu command list and the preference key list; embed it again after updating
- Trailing spaces are part of some IDs (e.g. `'Live PSAdapter_plugin_Ct  '`). Do not remove them from the output code
- A shortcut file is created when you save a set in Keyboard Shortcuts. Without one, Recheck cannot compare menu commands and tools
- Absence from the shortcut or preference files does not prove removal: commands outside the shortcut list and settings never changed are not written to those files
- Copy code by selecting it in the code field (Illustrator scripts cannot send text to the clipboard directly)

### Changelog

- v1.0.0 (20260927) : Initial release
