# ConvertAreaAndPointType

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertAreaAndPointType.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ConvertAreaAndPointType.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that converts point text / path text → area type (forward) or area type → point text (reverse), choosing the direction automatically from the selection.
- The forward conversion keeps the original appearance; the reverse conversion can keep the original frame size.
- Both directions share one dialog (title "Convert Text") showing a summary of the current selection and the Keep / Don't-keep style choice.

### Main Features

- **Selection**: summarizes the current selection by type, e.g. "Point text ×2"
- **Style: Keep** (default) / **Don't keep** — the meaning differs per direction, clarified by tooltips
  - Forward · Keep: register the source text's appearance (fill / stroke / effects) as a temporary graphic style and apply it after conversion (the temp style is removed afterwards)
  - Forward · Don't keep: transfer only the text, applying no graphic style
  - Reverse · Keep: leave the converted look as-is; no extra processing
  - Reverse · Don't keep: add two fills plus an absolute-size (width A × height B) "Convert to Shape: Rectangle" effect as a button-like background
- The frame size is based on the real size measured via duplicate → expand appearance → create outlines
- The forward conversion centers the text both horizontally and vertically, and carries over auto-kerning and mojikumi settings
- Path text is first detached into point text (keeping per-character attributes and justification), then converted through the same path
- Vertical centering and graphic-style registration use temporary dynamic actions that are unloaded on exit, so nothing is left behind in the Actions panel
- If nothing could be converted, the script alerts with the reason instead of exiting silently
- Automatic Japanese / English UI

### Workflow

1. If the selection contains area type, run the reverse conversion; otherwise run the forward conversion
2. Review the selection summary in the shared dialog and choose whether to keep the style (Cancel aborts)
3. Forward: create a rectangle at the measured size → convert it to area type → transfer contents, font, kerning, and mojikumi, then center (applying the temp graphic style when "Keep")
4. Reverse: read the original frame size → convert to point text (recovering the converted frame by a temporary marker name) → add the rectangle background when "Don't keep"
5. Select the resulting objects and finish

### Not Supported

- No open document
- Nothing selected, or a selection containing no text
- Non-text objects (counted as "Non-text" in the summary but not converted)
- With a multiple selection, each item is converted individually

### Update History

- v1.0.0 (20260702): Initial release. Converts point / path text ⇔ area type depending on the selection. A shared "Convert Text" dialog shows the current selection and a Keep / Don't-keep style choice. Forward: area type at the appearance-expanded measured size. Reverse: point text (Don't keep adds two fills plus an absolute-size rectangle background). Vertical centering and graphic-style registration use temporary dynamic actions
