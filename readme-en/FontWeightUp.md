# Switch the selected text to the next heavier weight

[![Direct](https://img.shields.io/badge/Direct%20Link-FontWeightUp.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FontWeightUp.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontWeightUp.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that looks up the font of the selected text and applies the next heavier weight in the same family
- To compare weights with samples and choose one, use [FontWeightPicker](FontWeightPicker.md)

### Main Features

- Switches to the next heavier style within the same family and series (italic, widths such as Condensed, optical sizes such as Display)
- Mixed fonts are each stepped up by one weight
- Weights are judged the same way as [TypefaceSampler](TypefaceSampler.md)
  - Numbers such as W3, W6 and 45 Light
  - Abbreviations such as L / R / M / DB / B / EB / H / U
  - English weight names from Hair to Ultra Black (SemiBold, XBold, etc.)

### Usage

1. Select text frames or groups (or select characters with the Type tool)
2. Run the script

### Notes

- Fonts listed as separate families are treated as one family when registered, thin to heavy, in `FAMILY_GROUPS` at the top of the script (sw-L / sw-R / sw-B / sw-H are registered)
- Composite fonts and the heaviest weight of a family are left unchanged; those fonts are listed at the end
- Characters are rewritten one by one, so long text may take a while

### Article

- [DTP Transit Annex on note (Japanese)](https://note.com/dtp_tranist/n/n255437cfdba0)

### Update History

- v1.0.0 (20260927): Initial version
- v1.0.1 (20260928): sw-L / sw-R / sw-B / sw-H are treated as one family

### Script info

- Version: v1.0.1
