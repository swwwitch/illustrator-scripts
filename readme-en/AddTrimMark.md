# Create trim marks

[![Direct](https://img.shields.io/badge/Direct%20Link-AddTrimMark.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AddTrimMark.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMark.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script to create trim marks for an artboard or selected objects.
- Trim marks are placed on a dedicated "Trim" layer, and a guide of the original object is also automatically created.

### Main Features

- Create trim marks based on selected object shape if available
- If no selection, use the entire artboard rectangle as base
- Move trim marks to a "Trim" layer and lock it automatically
- Duplicate the original object and convert to guide

### Process Flow

1. Use selected object if available, otherwise create rectangle from artboard
2. Duplicate target object, remove fill and stroke
3. Execute trim mark creation menu, then delete duplicate object
4. Move trim marks to "Trim" layer and lock it
5. Duplicate the original object and convert to guide

### note

- [【Illustrator】現在のアートボードにトンボを作成する｜DTP Transit 別館](https://note.com/dtp_tranist/n/n40e3e39cf9f2)

### Update History

- v1.2.0 (20260401): Initial release
- v1.2.1 (20260831): Restore the preference, active layer, and layer visibility after the run; make the rectangle tolerance check effective and reject degenerate zero-area paths; stop accumulating duplicate layers on the All Artboards run
- v1.2.3 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.2.4 (20260928): The button row is now built with the shared part
- v1.2.5 (20260929): Dialog opacity changed to 98%
- v1.2.6 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.2.7 (20260930): Button rows with only right-side buttons are now centered