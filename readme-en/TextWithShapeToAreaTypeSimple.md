# TextWithShapeToAreaTypeSimple

[![Direct](https://img.shields.io/badge/Direct%20Link-TextWithShapeToAreaTypeSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextWithShapeToAreaTypeSimple.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextWithShapeToAreaTypeSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Convert point text or path text into area type while preserving its appearance.

Flow:

1. On launch, show the options dialog (size adjustment On/Off with width% / height%).
2. Before converting, register the selected text's appearance as a temporary graphic style.
3. Point / path text → area type at the measured real size (scaled by the width/height ratios when size adjustment is On), then apply the temp style to carry over the appearance (the temp style is removed afterwards).
4. When size adjustment is On, add a button background (two New Fills + a rectangle shape effect).
5. Justification and vertical alignment are always centered.

### Notes

- Frame size is based on the real size measured via duplicate → expand appearance → create outlines.
- Vertical centering and graphic-style registration use dynamic actions that are loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.
- If nothing could be converted, the script alerts instead of exiting silently.

### Supported selections

- Point text … measured directly and converted to area type
- Path text … first detached into point text internally (detachPathTextToPointText), then converted through the same path
- With a multiple selection, each is converted individually

### How the options behave

| | Size ratio | Button background (2 fills + rectangle effect) |
|---|---|---|
| On | Yes (e.g. W×1.2 / H×1.6) | Yes |
| Off | 1× | No |

- The button background adds two fills via New Fill and reshapes them with the rectangle shape effect. Fill colors are not set (the added fills keep their defaults).

### Script info

- Version: v1.3.0
