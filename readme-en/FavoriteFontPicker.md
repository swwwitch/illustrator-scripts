# Pick from your favorite fonts and apply one

[![Direct](https://img.shields.io/badge/Direct%20Link-FavoriteFontPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FavoriteFontPicker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteFontPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that lists the fonts you use often, narrowed by category and standard, and applies the one you pick to the selected text
- Of the same typeface in several standards (Pr6N, Pr5, etc.), only the highest-priority checked standard stays in the list
- Chinese, Hangul, Thai and other multilingual fonts are left out of the list by default
- The font you pick in the list is previewed on the selected text right away
- Adapted from "よく使うフォントパネル" (Favorite Font Panel, favoriteFont_AI.jsx v1.1.2, MIT License) by KOUJI & 相棒（Gem）

### Main Features

**List**

- Two columns: filter checkboxes on the left, the font list on the right
- Each family is a header with its weights indented below (single-style families and composite fonts take one row)
- Show PostScript names: one row per font by PostScript name (such as RyuminPr6N-Light)
- On opening, the font of the selected text is chosen in the list

**Filtering**

- Category: the checked ones are combined
  - Custom: fonts added with Add to Custom, shown without thinning out standards or exclusions (Exclude non-Japanese included)
  - Document: fonts recorded in `_ProjectFonts.txt` next to the document, shown without thinning out standards or exclusions (Exclude non-Japanese included)
  - Morisawa / TB (TypeBank) / FOT (Fontworks) / Hiragino: matched by the start of the family or PostScript name
  - Adobe (Kozuka, Source Han, Ryo): Adobe's own Japanese typefaces (Kazuraki included). Fonts delivered by Adobe Fonts cannot be told by name, so they are not included
  - Composite fonts: those made with Type > Composite Fonts, not filtered or thinned out by standard
- Show all: ignores Category, standard thinning and `EXCLUDE_FONTS`, and lists every font (Standard and Exclude non-Japanese still apply, composite fonts included)
- Standard: plain standards (Std, Pro, Pr5, Pr6) on the left, N variants on the right. Shows only the checked standards (and No standard); for each typeface the highest-priority checked standard is kept (uncheck Pr6N and Pr6 or Pr5 stays)
- Filter by standard: when off, the Standard checkboxes are ignored (each typeface keeps its highest-priority standard of all). On by default
- Option (Alt)-click a Category or Standard checkbox to switch between "only this one" and "all on"
- Filter: substring match on the family, style or PostScript name
- Exclude non-Japanese: leaves Chinese / Hangul / Thai / Other multilingual fonts out of the list (all on by default; applies under Show all too, while listed Custom and Document fonts stay). The language is judged by the font name

**Managing fonts**

- Add to Custom / Remove from Custom: puts the family of the chosen font in or out of Custom. Removing a name that also matches other fonts (such as "DIN") asks first
- Record used fonts in folder: adds the fonts used in the document to `_ProjectFonts.txt` (the same file as the InDesign version)
- Rescan: reads the installed fonts again

**Applying**

- Preview: applies the chosen font to the selected text for a look (on by default; Cancel restores it)
- Double-click, press Enter or click Apply to apply and close
- The checkboxes (Exclude non-Japanese included), the Custom fonts and the filter text are restored next time

### Usage

1. Select a text frame or a group (or some characters with the Type tool)
2. Run the script
3. Narrow the list with the checkboxes on the left and the Filter field
4. Pick a font to see the preview, then double-click, press Enter or click Apply

The Filter field updates the list as you type while the result is 800 fonts or fewer; above that it shows only the count until you press Enter (change the limit with `LIVE_SEARCH_MAX_FONTS`).

### Options (user settings at the top of the script)

| Setting | Description | Default |
|---|---|---|
| CUSTOM_FONTS | Initial Custom fonts; later changed in the dialog and saved | Graphik, DIN |
| EXCLUDE_FONTS | Fonts left out of the list (substring match; ignored by Show all, Custom and Document) | NT |
| RESCUE_FONTS | Fonts kept even when EXCLUDE_FONTS matches them (substring match) | none |
| PRIORITY_PREFIXES | Prefix priority (first wins) | A-OTF, A P-OTF, AP-OTF, G-OTF, U-OTF |
| PRIORITY_SUFFIXES | Standard priority; also listed as Standard checkboxes | Pr6N, Pr6, Pr5N, Pr5, ProN, Pro, StdN, Std |
| FOUNDRY_FILTERS | Foundry categories (prefix match on the family or PostScript name) | Morisawa, TB, FOT, Hiragino, Adobe |
| LANGUAGE_GROUPS | Languages under Exclude non-Japanese and the name rules that detect them (prefixes, regular expressions) | Chinese, Hangul, Thai, Other multilingual |
| LIVE_SEARCH_MAX_FONTS | Max result count updated on every keystroke | 800 |

Custom and Document names match whole words ("Pr5" does not match "Pr5N"; "DIN" matches "DIN 2014").

### Notes

- On the first run, and whenever the installed fonts change (judged by the total and 32 names sampled at even intervals), the font information is read before the dialog opens (with a progress bar). If a change slips through, click Rescan. The cache is `~/Library/Application Support/illustrator-scripts/FavoriteFontPicker-fonts.txt`
- The font is applied to the text that was selected when the dialog opened. The selection cannot change while the dialog is open, so it closes after applying
- The preview duplicates the text and hides the original for a while. Characters selected with the Type tool and threaded text are not previewed (Apply still applies to them)
- Missing fonts are not listed
- Locked or hidden text is left as is
- Record used fonts in folder needs a document that has been saved

### Article

- [DTP Transit 別館｜note](https://note.com/dtp_tranist/n/ncf9ff6feebf0)

### Update History

- v1.0.1 (20261002) : Bug fixes
  - Record used fonts in folder no longer records fonts tried in the preview
  - Clicking a family header no longer selects the family above
  - Custom matches by prefix again ("DIN" also shows DINPro and DINOT)
  - Hiragino Sans CNS now counts as Chinese under Exclude non-Japanese
  - Options > Exclude is now a single panel titled Exclude non-Japanese
- v1.0.0 (20261001) : Initial version, adapted from "よく使うフォントパネル" v1.1.2
  - Two-column dialog; a list with weights under family headers; Show PostScript names
  - Standard is now a set of checkboxes that also works without Show all; Filter by standard; Option-click switching
  - Filter field, adding and removing Custom fonts, preview, the Apply button, preselecting the current font, saved state
  - Hiragino, Adobe and Composite fonts categories, and Options to exclude languages
  - Rescans automatically when the installed fonts change
  - Morisawa detection no longer picks up every name containing "ud", and Document names match whole words. Document fonts are no longer hidden by the standard thinning
