# ImportAndApplyGraphicStyle

[![Direct](https://img.shields.io/badge/Direct%20Link-ImportAndApplyGraphicStyle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ImportAndApplyGraphicStyle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportAndApplyGraphicStyle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Convert point text, path text, or text + shape into area type while preserving appearance.

Flow:

1. On launch, show the resident palette (size adjustment On/Off with width% / height%, and a graphic-style choice: original appearance / a loaded style). Run via the "Convert" button; the palette stays open.
2. Before converting, prepare the graphic style to apply ("original appearance" = register the selected text's appearance temporarily; a loaded style = import from the AI file picked via "Load").
3. Point / path text → area type at the measured real size (scaled by the width/height ratios when size adjustment is On).
   Text + rectangle → duplicate the rectangle into an area-type frame and fill it with the text.
4. Button background (text-only "On" / text + rectangle; only when no external style is used) → add two fills via New Fill + a rectangle shape effect.
5. Apply the prepared graphic style to the converted area type (remove only the temp "original appearance" style; keep imported external styles).
6. Justification and vertical alignment are always centered.

### Notes

- Only a rectangle (axis-aligned, straight corners) can serve as the text + shape frame. If the only shape selected is a non-rectangle closed path, the script alerts and aborts.
- Text + shape processes a single pair only (the first text + the first rectangle).
- Frame size is based on the real size measured via duplicate → expand appearance → create outlines.
- Vertical centering and graphic-style registration use dynamic actions that are loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.

### Supported selections

① Text only
- Point text … measured directly and converted to area type
- Path text … first detached into point text internally (workerDetachPathText), then converted through the same path
- With a multiple selection, each is converted individually

② Text + rectangle
- Text (point / path) + a rectangle (axis-aligned, straight corners)
- The rectangle is duplicated into the area-type frame and filled with the text
- Only a single pair is processed (the first text + the first rectangle)

### How the options behave

| | Size ratio | Button background (2 fills + rectangle effect) |
|---|---|---|
| Text only · On | Yes (e.g. W×1.2 / H×1.6) | Yes |
| Text only · Off | 1× | No |
| Text + rectangle | — (rectangle size wins) | Yes |

- The button background adds two fills via New Fill and reshapes them with the rectangle shape effect. Fill colors are not set (the added fills keep their defaults).
- When a loaded style is chosen, the style defines the appearance, so no button background is added. If not registered in the current document, it is imported from the remembered AI file and applied. The source file and its style names are remembered in Folder.userData (styles_for_TextWithShapeToAreaType.txt).

### Script info

- Version: v1.3.0
