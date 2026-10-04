# Register selected objects as symbols

[![Direct](https://img.shields.io/badge/Direct%20Link-SymbolizeEach.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/SymbolizeEach.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeEach.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Registers the selected objects as symbols, one by one or as a single symbol, and replaces the originals with symbol instances.
Names are assigned automatically from the text contents, layer name, note, or a sequence number, or confirmed in Illustrator's native New Symbol dialog.

### Features

- Registers the whole selection as one symbol, or each object as its own symbol
- In automatic registration, the symbol name is chosen in this order:
  1. Text contents (a text frame, or the first text frame found inside a group)
  2. Layer name (default names such as "Layer 1" are skipped)
  3. Note (Attributes panel)
  4. Prefix + sequence number
- Appends `_2`, `_3` … when a symbol with the same name already exists
- Places each instance at the original's position (top-left) and stacking order
- Linked images can be embedded and registered, or excluded
- Existing symbol instances in the selection are left as they are
- Shows the number of symbols created, items skipped, and failures (with details) when done

### Usage

1. Select the objects you want to register as symbols.
2. Run the script.
3. Set the options in the dialog and click OK.

### Options

- Selection handling: Create one symbol from selection / Create symbols per object
- Registration method
  - Confirm with native dialog: opens Illustrator's New Symbol dialog for each target so you set the name and registration point each time
  - Register automatically: registers without confirmation, using the Symbol name and Registration point settings below
- Symbol name (automatic registration only)
  - Prefix: used when no name can be taken from text contents, the layer name, or the note (`Symbol_` when left blank)
  - Sequence: number of digits after the prefix (0 / 00 / 000)
  - Use text contents as symbol name: when off, naming starts from the layer name
- Registration point: chosen on a 3×3 grid (automatic registration only)
- Linked images
  - Ignore: keeps them selected but excludes them from registration
  - Embed and register: embeds them before creating the symbol

### Notes

- Choosing "Create one symbol from selection" fixes the registration method to "Confirm with native dialog".
- With "Confirm with native dialog", the view zooms to each target and returns to the original view when finished.
- Sequence numbers follow the order within the selection, so excluded items such as existing symbol instances leave gaps in the numbering.
- The default prefix, sequence digits, and initial selections can be changed in the "User settings" block at the top of the script.
- Use [SymbolizeAndReplace](SymbolizeAndReplace.md) when matching items elsewhere in the document should be replaced too.

### Article

https://note.com/dtp_tranist/n/nce9ec30232a0 (Japanese)

### Update History

- v1.0.11 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.0.10 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.0.9 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.0.8 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.0.7 (2026-09-30) Button rows with only right-side buttons are now centered
- v1.0.6 (2026-09-30) Fixed an error when running with characters selected by the Type tool
- v1.0.5 (2026-09-29) Dialog opacity changed to 98%
- v1.0.4 (2026-09-28) Clip groups are now measured by their mask
- v1.0.4 (2026-09-28) The 3×3 reference point picker now uses the shared part
- v1.0.4 (2026-09-28) The button row is now built with the shared part
- v1.0.3 (2026-09-28) The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.2 (2026-09-15) Replaced the registration-point radio buttons with a 3×3 anchor widget, tidied the field labels, and reorganized internal naming and functions
- v1.0.1
