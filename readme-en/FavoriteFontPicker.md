# Pick from your favorite fonts and apply one

[![Direct](https://img.shields.io/badge/Direct%20Link-FavoriteFontPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FavoriteFontPicker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteFontPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that lists only the fonts you use often and applies the one you pick to the selected text
- Of the same typeface in several standards (Pr6N, Pr5, etc.), only the highest-priority one stays in the list
- Adapted from "よく使うフォントパネル" (Favorite Font Panel, favoriteFont_AI.jsx v1.1.2, MIT License) by KOUJI & 相棒（Gem）

### Main Features

- Two columns: checkboxes on the left, the font list on the right
- Category: Custom / Document / Morisawa / TB (TypeBank) / FOT (Fontworks) (the checked ones are combined)
  - Custom: fonts listed in `CUSTOM_FONTS` at the top of the script (prefix match)
  - Document: fonts listed in `_ProjectFonts.txt` next to the document, shown as is without thinning out standards
- Standard: plain standards (Std, Pro, Pr5, Pr6) on the left, N variants on the right. Shows only the checked standards (and No standard). For each typeface, the highest-priority checked standard is kept (uncheck Pr6N and Pr6 or Pr5 stays)
- Show all: ignores Category, standard thinning and exclusions and lists every font (Standard still filters)
- Option (Alt)-click a Category or Standard checkbox to switch between "only this one" and "all on"
- Filter: substring match on the family, style or PostScript name
- Show PostScript names: lists PostScript names (such as RyuminPr6N-Light) and sorts by them
- Record used fonts in folder: adds the fonts used in the document to `_ProjectFonts.txt` (the same file as the InDesign version)
- Rescan: reads the installed fonts again
- The checkboxes and the filter text are restored next time

### Usage

1. Select a text frame or a group (or some characters with the Type tool)
2. Run the script
3. Double-click a font in the list, or select one and press Enter or click Apply

The Filter field updates the list as you type while the result is 800 fonts or fewer; above that it shows only the count until you press Enter (change the limit with `LIVE_SEARCH_MAX_FONTS`).

### Options (user settings at the top of the script)

| Setting | Description | Default |
|---|---|---|
| CUSTOM_FONTS | Fonts listed under Custom (prefix match) | Graphik, DIN |
| EXCLUDE_FONTS | Fonts left out of the list (substring match; ignored by Show all) | NT |
| RESCUE_FONTS | Fonts kept even when EXCLUDE_FONTS matches them (substring match) | none |
| PRIORITY_PREFIXES | Prefix priority (first wins) | A-OTF, A P-OTF, AP-OTF, G-OTF, U-OTF |
| PRIORITY_SUFFIXES | Standard priority; also listed as Standard checkboxes | Pr6N, Pr6, Pr5N, Pr5, ProN, Pro, StdN, Std |
| FOUNDRY_FILTERS | Foundry categories (prefix match on the family or PostScript name) | Morisawa, TB, FOT |

### Notes

- On the first run, and whenever the installed fonts change (judged by the total and 32 names sampled at even intervals), the font information is read before the dialog opens (with a progress bar). If a change slips through, click Rescan. Later runs use the cache (`~/Library/Application Support/illustrator-scripts/FavoriteFontPicker-fonts.txt`)
- Composite fonts and missing fonts are not listed
- Record used fonts in folder needs a document that has been saved
- Locked or hidden text is left as is

### Article

- [DTP Transit 別館｜note](https://note.com/dtp_tranist/n/ncf9ff6feebf0)

### Update History

- v1.0.0 (20261001) : Initial version, adapted from "よく使うフォントパネル" v1.1.2. Added the two-column dialog, Standard checkboxes (always available), the filter field, Show PostScript names, the Apply button, saved state, and an automatic rescan when the number of fonts changes. Morisawa detection no longer picks up every name containing "ud". Document fonts are no longer hidden by the standard thinning
