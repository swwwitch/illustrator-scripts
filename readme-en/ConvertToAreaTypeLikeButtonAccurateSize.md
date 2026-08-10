# ConvertToAreaTypeLikeButtonAccurateSize

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToAreaTypeLikeButtonAccurateSize.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ConvertToAreaTypeLikeButtonAccurateSize.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToAreaTypeLikeButtonAccurateSize.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Convert point text, path text, or text + shape into area type while preserving appearance.

Flow:

1. Before converting, register the selected text's appearance as a temporary graphic style (dynamic action).
2. Point / path text → area type at the measured real size.
   Text + shape → fill the shape (as area type) with the text.
3. Apply the registered graphic style to the converted area type, then remove the temporary style.
4. Justification and vertical alignment are always centered.

### Notes

- Frame size is based on the real size measured via duplicate → expand appearance → create outlines.
- Vertical centering and graphic-style registration use dynamic actions that are loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.

### Script info

- Version: v1.0.0
