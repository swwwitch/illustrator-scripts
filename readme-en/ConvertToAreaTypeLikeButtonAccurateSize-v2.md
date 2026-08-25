# ConvertToAreaTypeLikeButtonAccurateSize-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToAreaTypeLikeButtonAccurateSize--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ConvertToAreaTypeLikeButtonAccurateSize-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToAreaTypeLikeButtonAccurateSize-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Convert the selected text in either direction, based on the selection.

- Point / path text → area type, preserving appearance and real size.
- Area type → point text.

Flow (point / path text → area type):

1. Path text is first detached into point text (glyphs and attributes preserved).
2. Before converting, register the text's appearance as a temporary graphic style (dynamic action).
3. Build a rectangle frame at the real size measured via duplicate → expand appearance → create outlines, then convert it to area type.
4. Carry over the source text's contents, font, size, kerning, and mojikumi settings.
5. Apply the registered graphic style, then remove the temporary style.
6. Justification (horizontal) and text placement (vertical) are always centered.

Flow (area type → point text):

1. Read the area text's width and height (width A, height B).
2. Convert to point text via convertAreaObjectToPointObject().
3. Measure the text's real size (width D, height E) via duplicate → create outlines, then discard the copy.
4. Apply the Convert-to-Shape (Rectangle) effect with extra width = A − D and extra height = B − E, reproducing a rectangle at the original frame size.

### Notes

- If the selection contains area type, it is treated as the reverse conversion (→ point text); otherwise as the forward conversion (→ area type).
- Frame size is based on the real size measured via duplicate → expand appearance → create outlines.
- Vertical centering and graphic-style registration use dynamic actions that are loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.

### Script info

- Version: v1.0.0
