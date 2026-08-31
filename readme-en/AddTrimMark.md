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