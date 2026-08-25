# SymbolChangeAxis

[![Direct](https://img.shields.io/badge/Direct%20Link-SymbolChangeAxis.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/SymbolChangeAxis.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolChangeAxis.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Breaks the links of the selected symbol instances and turns them into regular editable objects.

Options are provided for tidying up the resulting hierarchy and names.

### Features

- Processes symbols anywhere in the selection, including inside groups
- Tidies the results of breaking static and dynamic symbols
- Carries the original symbol name over to the resulting items
- Optional name prefix
- Uses the text content as the name when the result is a single text frame
- Automatically releases groups that contain only one item

### Usage

1. Select the symbol instances whose links you want to break.
2. Run the script.
3. Set the naming and hierarchy options and click OK.

### Notes

- Use AiBreakLink.jsx when you only need a quick break-link.

### Update History

- v1.1.0
