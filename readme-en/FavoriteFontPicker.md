# Pick from your favorite fonts and apply one

[![Direct](https://img.shields.io/badge/Direct%20Link-FavoriteFontPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FavoriteFontPicker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteFontPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that lists the fonts you use often, narrowed by category and standard, and applies the one you pick to the selected text
- Of the same typeface in several standards (Pr6N, Pr5, etc.), only the highest-priority checked standard stays in the list
- Chinese, Korean, Thai and other non-Japanese fonts are left out of the list by default
- The font you pick in the list is previewed on the selected text right away
- Adapted from "よく使うフォントパネル" (Favorite Font Panel, favoriteFont_AI.jsx v1.1.2, MIT License) by KOUJI & 相棒（Gem）. Main changes:
  - The single-column dialog is now two columns: filters on the left, the list on the right
  - The tree view is now a list with weights under family headers
  - The standard dropdown, once usable only under Show all, is now the Standard checkboxes (they work without Show all too)
  - Hiragino, Adobe, Adobe Fonts and Composite fonts categories added; Morisawa detection now matches by prefix
  - Japanese fonts only added (the languages to leave out — Chinese, Korean, Thai, Other — are chosen in Preferences)
  - Fonts whose PostScript names contain non-ASCII characters, and composite fonts, are now listed (they were left out)
  - Document font names now match whole words instead of prefixes ("Pr5" does not match "Pr5N")
  - Custom became five Custom sets, filled with the buttons below the list
  - Preferences… added, gathering Rescan, the languages to leave out, and Custom set export and import
  - Filter field, font count and Show PostScript names added
  - Preview on the selected text, plus Apply and Cancel buttons (double-click or Enter in the list also applies)
  - The current font of the selected text is preselected on open
  - Checkboxes and the filter text are restored next time
  - A progress bar shows while font information is read; the cache is checked and reread automatically when fonts change
  - The "extract from the current document" button removed (reading _ProjectFonts.txt stays)
  - English UI

<img alt="The Favorite Fonts dialog" src="../png/ss-1148-1320-144-20261002-070229-s.png" width="50%" />

### Main Features

**List**

- Two columns: filter checkboxes on the left, the font list on the right
- Each family is a header with its weights indented below (single-style families and composite fonts take one row)
- Show PostScript names: one row per font by PostScript name (such as RyuminPr6N-Light)
- On opening, the font of the selected text is chosen in the list

**Filtering**

- Category: the checked ones are combined
  - Custom sets 1–5: fonts added with the numbered buttons below the list; the checked sets are combined, shown without thinning out standards or exclusions (Japanese fonts only included)
  - Document fonts: fonts listed in `_ProjectFonts.txt` next to the document (the file made by the InDesign version), shown without thinning out standards or exclusions (Japanese fonts only included)
  - Composite fonts: those made with Type > Composite Fonts, not filtered or thinned out by standard
  - Morisawa / TB (TypeBank) / FOT (Fontworks) / Hiragino: matched by the start of the family or PostScript name
  - Adobe: Adobe's own Japanese typefaces (Kozuka, Source Han, Ryo, Ten Mincho, Momochidori and Kazuraki)
  - Adobe Fonts: fonts activated through Creative Cloud, judged by the family names in the Creative Cloud sync state (`entitlements.xml`)
- Show all: ignores Category, standard thinning and `EXCLUDE_FONTS`, and lists every font (Standard and Japanese fonts only still apply, composite fonts included)
- Standard: plain standards (Std, Pro, Pr5, Pr6) on the left, N variants on the right. Shows only the checked standards (and No standard); for each typeface the highest-priority checked standard is kept (uncheck Pr6N and Pr6 or Pr5 stays)
- Filter by standard: when off, the Standard checkboxes are ignored (each typeface keeps its highest-priority standard of all). On by default
- Option (Alt)-click a Category, Standard or Preferences language checkbox to switch, panel by panel, between "only this one" and "all on"
- Filter: substring match on the family, style or PostScript name; the circled × button on the right clears it
- Japanese fonts only: leaves fonts in the languages chosen in Preferences (Chinese / Korean / Thai / Other; all by default) out of the list (on by default; applies under Show all too, while listed Custom set and Document fonts stay). The language is judged by the font name

**Managing fonts**

- Numbered buttons (1–5): put the family of the chosen font in or out of that Custom set (for the sets it is in, the whole icon turns dark and the bookmark is filled). Removing a name that also matches other fonts (such as "DIN") asks first
- Preferences… (bottom left):
  - Rescan: reads the installed fonts again
  - Choose the languages left out by Japanese fonts only
  - Show each Custom set, and export it to or import it from a text file with one name per line (importing adds to the set and skips duplicates)

**Applying**

- Preview on selected text (bottom left, next to Preferences…): applies the chosen font to the selected text for a look (on by default; Cancel restores it)
- Double-click, press Enter or click Apply to apply and close
- The checkboxes (Japanese fonts only included), the Custom set fonts and the filter text are restored next time

### Usage

1. Select a text frame or a group (or some characters with the Type tool)
2. Run the script
3. Narrow the list with the checkboxes on the left and the Filter field
4. Pick a font to see the preview, then double-click, press Enter or click Apply

To keep fonts you use often, choose one in the list and click a numbered button below it to put it in that Custom set. Next time, check that number under Category to bring them back.

The Filter field updates the list as you type while the result is 800 fonts or fewer; above that it shows only the count until you press Enter (change the limit with `LIVE_SEARCH_MAX_FONTS`).

### Options (user settings at the top of the script)

| Setting | Description | Default |
|---|---|---|
| CUSTOM_FONTS | Initial fonts of Custom set 1; later changed in the dialog and saved | Graphik, DIN |
| CUSTOM_SET_COUNT | Number of Custom sets | 5 |
| EXCLUDE_FONTS | Fonts left out of the list (substring match; ignored by Show all, Custom and Document fonts) | NT |
| RESCUE_FONTS | Fonts kept even when EXCLUDE_FONTS matches them (substring match) | none |
| PRIORITY_PREFIXES | Prefix priority (first wins) | A-OTF, A P-OTF, AP-OTF, G-OTF, U-OTF |
| PRIORITY_SUFFIXES | Standard priority; also listed as Standard checkboxes | Pr6N, Pr6, Pr5N, Pr5, ProN, Pro, StdN, Std |
| FOUNDRY_FILTERS | Foundry categories (prefix match on the family or PostScript name) | Morisawa, TB, FOT, Hiragino, Adobe, Adobe Fonts |
| LANGUAGE_GROUPS | Languages left out by Japanese fonts only and the name rules that detect them (prefixes, regular expressions) | Chinese, Korean, Thai, Other |
| LIVE_SEARCH_MAX_FONTS | Max result count updated on every keystroke | 800 |

Custom names match by prefix ("DIN" matches "DIN 2014" and "DINPro"). Document font names match whole words ("Pr5" does not match "Pr5N").

### Notes

- On the first run, and whenever the installed fonts change (judged by the total and 32 names sampled at even intervals), the font information is read before the dialog opens (with a progress bar). If a change slips through, click Rescan in Preferences…. The cache is `~/Library/Application Support/illustrator-scripts/FavoriteFontPicker-fonts.txt`
- The font is applied to the text that was selected when the dialog opened. The selection cannot change while the dialog is open, so it closes after applying
- The preview duplicates the text and hides the original for a while. Characters selected with the Type tool and threaded text are not previewed (Apply still applies to them)
- Missing fonts are not listed
- Locked or hidden text is left as is

### Article

- [DTP Transit 別館｜note](https://note.com/dtp_tranist/n/ncf9ff6feebf0)

### Update History

- v1.1.0 (20261002) : New features and bug fixes
  - Added Preferences…, which now holds Rescan and the languages to leave out, plus showing, exporting and importing Custom sets. Exclude non-Japanese became Japanese fonts only in the main window
  - Clicking a family header no longer selects the family above
  - Custom matches by prefix again ("DIN" also shows DINPro and DINOT)
  - Hiragino Sans CNS now counts as Chinese under Exclude non-Japanese
  - Options > Exclude is now a single panel titled Exclude non-Japanese
  - UI wording revised: Document → Document fonts, Hangul / Other multilingual → Korean / Other; tooltips added to the foundry categories and Apply
  - Composite fonts now follows Document fonts in Category, with a divider before the foundry categories
  - Added a clear button to the right of Filter, and more space below the font list
  - Option-click switching now works in Exclude non-Japanese too, which is laid out in two columns
  - Show all moved above Category, and Filter by standard to the top of the Standard panel
  - Custom is now five Custom sets (the old Custom carries over to set 1), with five numbered buttons below the list
  - Checkboxes respond faster
  - Added Adobe Fonts to Category (above the foundry divider); Ten Mincho and Momochidori now count as Adobe
  - Removed Record used fonts in folder. The count and Show PostScript names now sit above the list, Preview moved next to Rescan as Preview on selected text, Add to Custom and Remove from Custom are one bookmark toggle button, and the list grows when the dialog is made taller
- v1.0.0 (20261001) : Initial version, adapted from "よく使うフォントパネル" v1.1.2
  - Two-column dialog; a list with weights under family headers; Show PostScript names
  - Standard is now a set of checkboxes that also works without Show all; Filter by standard; Option-click switching
  - Filter field, adding and removing Custom fonts, preview, the Apply button, preselecting the current font, saved state
  - Hiragino, Adobe and Composite fonts categories, and Options to exclude languages
  - Rescans automatically when the installed fonts change
  - Morisawa detection no longer picks up every name containing "ud", and Document names match whole words. Document fonts are no longer hidden by the standard thinning
