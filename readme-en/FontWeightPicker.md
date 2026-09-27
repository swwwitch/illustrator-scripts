# Compare weights of similar fonts with samples and apply one

[![Direct](https://img.shields.io/badge/Direct%20Link-FontWeightPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FontWeightPicker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontWeightPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that lists the weights of the selected text's font and of families with similar names, and applies the font you choose while you compare samples
- To step up one weight right away, use [FontWeightUp](FontWeightUp.md)

### Main Features

- Font list on the left, weights on the right
  - The right side lists the weights in the same series (italic, widths such as Condensed), lightest first
  - Weights used in the selected text are dimmed and cannot be chosen
  - Choosing a radio button previews it on the selected text right away
- Related fonts: families sharing the first word (Bebas Neue and Bebas Kai for Bebas Neue Pro) or containing the core name (Shin Go ProN, UD Shin Go, etc. for A-OTF Shin Go Pr6N)
- Weight samples: a black sample of each weight in a column to the right of the selected text, with the weight name beside it (the current weight level with the text, lighter above, heavier below)
- Cover background in white: a translucent white rectangle behind the samples fades other objects; a black copy of the selected text sits on top
- Fit to window: zooms so the selected text and the samples fit
- Weights are judged the same way as [TypefaceSampler](TypefaceSampler.md)

### Usage

1. Select text frames or groups (or select characters with the Type tool)
2. Run the script
3. Choose a font on the left and a weight on the right, then click OK (Cancel restores the original fonts and view)

### Options

| Option | Description | Default |
|---|---|---|
| Show Related Fonts | When off, only the selected family is listed | On |
| Show PostScript Names | Shows PostScript names (such as BebasNeuePro) in the font list | Off |
| Show Weight Samples on the Right | Places samples and weight names on a temporary layer | On |
| Cover Background in White [%] | Opacity of the white rectangle | On, 85% |
| Fit to Window [%] | Size relative to the window | On, 65% |

### Notes

- The list is based on the font of the first character; the chosen font is applied to every selected character
- Samples and the white rectangle are created on the temporary layer "// weight-preview" and removed when the dialog closes
- Composite fonts are not listed

### Update History

- v1.0.0 (20260928): Initial version

### Script info

- Version: v1.0.0
