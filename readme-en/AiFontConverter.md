# AiFontConverter

[![Direct](https://img.shields.io/badge/Direct%20Link-AiFontConverter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/AiFontConverter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiFontConverter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Switches font variants (character set, P, UD, N, NT, weight) in bulk.

- Target the selected objects, the whole document, or the active artboard
- Conversion settings sit in three columns (left: character set, center: N and NT, right: UD and P); each item toggles between keep / off / on
- Character sets are Std / Pro / Pr5 / Pr6, with N toggled on a separate axis
- Shin Go and Shin Go NT are switched through the NT setting
- Max and MaxN presets move fonts to the richest available character set (MaxN includes N)
- G-OTF Gakusan faces (Jokai / Gakusan / K) can be merged into A-OTF with a checkbox
- Special series such as A1 Mincho are mapped by equivalent weight (A-OTF A1 Mincho Std B = A P-OTF A1 Mincho StdN R)
- CID fonts and the pre-run confirmation are controlled by switches at the top of the script (hidden from the UI)
- Changes are applied per textRange for speed
- The convertible families are generated from a font database (`FONT_FAMILIES`), covering the core Morisawa faces plus the Tsukushi series, UD faces and Fontworks-derived designs (Cezanne, Matisse, Rodin, and so on)

### References

https://sttk3.com/blog/tips/illustrator/unify-character-set.html

### Update History

- v1.0.0: Initial release
- v1.0.1: Show Japanese font names in the confirmation dialog, order entries by canvas position (top to bottom), align the arrows by fixing the "before" column width, drop the version from the title, and adjust the side margins
- v1.1.0: Expanded the convertible families (Tsukushi series, UD faces, Fontworks-derived designs, and so on)
- v1.1.1: Added AXIS (Type Project) faces, with dedicated handling that preserves the width (Basic/Cond/Comp) and Joyo and switches only N and Std/Pro; the Max and MaxN presets always move them to ProN

### Article

https://note.com/dtp_tranist/n/n261c771b4b41

### Script info

- Version: v1.1.1
