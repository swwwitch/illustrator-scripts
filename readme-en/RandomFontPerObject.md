# Apply a random font to each object

[![Direct](https://img.shields.io/badge/Direct%20Link-RandomFontPerObject.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/RandomFontPerObject.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RandomFontPerObject.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Applies a random font to each selected text object. All the text in an object gets the same font. This is the font swap from [Auto Touch Type Tool](AutoTouchType.md), extracted and changed from per character to per object. The same feature is also in [the script that randomizes position, scale, rotation and more](RandomizeObjects.md).

### Main Features

- Picks one font per object. Text inside groups is included.
- **Japanese only** limits the candidates to Japanese fonts. It starts on when the selected text contains Japanese.
- **Random** draws different fonts with the same settings.

### Usage

1. Select text objects.
2. Run the script. The fonts change as soon as the dialog opens.
3. Adjust with **Random** and **Japanese only**.
4. Click **OK** to confirm, or **Cancel** to restore the original fonts.

### Notes

- When characters are selected with the Type tool, that range counts as one object.
- Locked and hidden text is skipped.
- Japanese fonts are detected from the font name.

### Update History

- v1.0.0 (2026-11-07): Initial release
- v1.0.1 (2026-11-08): Japanese font detection no longer counts Fontworks Chinese and Korean fonts, Fontworks non-Japanese scripts such as FWArabic, Korean fonts such as BM Dohyeon (배달의민족 도현), Fontworks emoji, kana-only fonts without kanji, Chinese fonts such as PingFang, LiHei and Hiragino Sans GB, Shizuka Gothic (靜呉竹), Tsukimi, or fonts whose names contain Thai or 體
