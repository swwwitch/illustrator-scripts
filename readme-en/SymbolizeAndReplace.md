# Make a symbol and replace matching objects with it

[![Direct](https://img.shields.io/badge/Direct%20Link-SymbolizeAndReplace.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/SymbolizeAndReplace.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeAndReplace.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Converts the selected object into a symbol and replaces matching items in the document with instances of that symbol.

### Features

- Set the symbol name and registration point in a dialog (empty names and names of existing symbols are rejected)
- Pick the registration point by clicking a widget of nine squares
- What gets replaced depends on the selection
  - A text frame: text frames with the same font, style, and contents. The symbol name starts as that text (line breaks become spaces). Turn on "Include different font sizes" to ignore size differences
  - Multiple groups: every selected group is replaced with one symbol (the frontmost group becomes the symbol's artwork)
  - Any other object: similar objects found with Global Edit
- Replaces each target with a symbol instance, aligned by the chosen registration point
- Leaves the new symbol instances selected after replacement
- Reports how many items could not be replaced (locked, hidden, and so on)

### How to use

1. Select the object to base the symbol on (a text frame, a single object, or several groups)
2. Run the script and set the symbol name and registration point
3. Click OK to create the symbol and replace the matching objects

### Notes

- A multiple selection runs only when every item is a group
- When nothing matches besides the original, the created symbol is removed and the script stops
- Does nothing while characters are selected with the Type tool
- The symbol type (dynamic or static) cannot be set

### Article

https://note.com/dtp_tranist/n/n650a4b91329d

### Update History

- v1.0.2 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.3 (2026-09-28): Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure. The button row is now built with the shared part
- v1.0.4 (2026-09-29): Dialog opacity changed to 98%
- v1.0.5 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.0.6 (2026-09-30): Button rows with only right-side buttons are now centered
- v1.0.7 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.0.8 (2026-10-01): The OK button label now comes from the label definitions. Unified the window and panel margins and spacing with the shared layout part
- v1.0.9 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.0.10 (2026-10-01): The registration point is now picked with the shared anchor widget (nine drawn squares) instead of nine radio buttons. The panel is now titled "Registration Point", and alerts appear only in the UI language. Running with characters selected by the Type tool now does nothing

### Script info

- Version: v1.0.10
