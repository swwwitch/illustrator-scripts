# Search and show panels

[![Direct](https://img.shields.io/badge/Direct%20Link-AiPanelLauncher.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiPanelLauncher.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiPanelLauncher.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README-en.md)

---

Lists the panels of the Window menu in a single list, including those in the Type submenu. Filter it with the search field and show the panels you pick. The Illustrator counterpart of [IdPanelLauncher](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdPanelLauncher.md) for InDesign.

### Features

- Shows the panels of the Window menu in a two-column list: Name | Category
  - Items in the Type submenu (Character, Paragraph, OpenType, etc.) are expanded and show "Type" as their category
  - Panels your version of Illustrator does not have yet (Retype, Concept to Vector, etc.) are left out
- Filter by name and category with the search field (as you type)
  - Separate words with spaces to match all of them
  - Case-insensitive. Japanese hiragana is matched as katakana, and kanji names also match their reading
  - The (×) button next to the field clears the search
- Switch the sort order with radio buttons
  - **Default**: the order of the Window menu (Type items stay where the Type submenu is)
  - **Alphabetical**: by panel name
  - The selected panels stay selected when you re-sort
- Filtering selects the first panel, so Enter shows it right away
- Pick several panels to show them at once
- The dialog title shows the number of panels listed, and the dialog opens where you left it

### Usage

1. Run the script
2. Type part of a panel name in the search field
3. Show it in one of these ways
   - Enter / Return
   - Double-click the item in the list
   - Click Show

Press Down in the search field to move to the list. Cmd (Ctrl) / Shift-click to pick more than one.

### Notes

- Scripts cannot read Illustrator's menus, so the panel list (names, menu command IDs, minimum versions) is kept in a table in the script (PANEL_TABLE), matching the Window menu of Illustrator 2026 (30.8). Extension and plug-in panels and the libraries are not listed.
- The Blend panel (30.7 and later) is not listed because its menu command ID is unknown.
- Showing a panel is the same as choosing its Window menu item (`app.executeMenuCommand()`). Illustrator's menu items toggle the panel and scripts cannot read whether it is open, so picking an open panel may close it.
- Panels are shown after the dialog closes.
- In the English UI the Window menu is alphabetical, so Default gives the same order as Alphabetical.
- The script uses `#targetengine` because it is meant to be run often. Assigning a keyboard shortcut is handy.
- Esc closes the dialog.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/misc/AiPanelLauncher.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-05 |
| Last updated | 2026-11-05 |
| Tested with | Illustrator 2026 (30.8) / macOS |

### Update History

- v1.0.0 (2026-11-05): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
