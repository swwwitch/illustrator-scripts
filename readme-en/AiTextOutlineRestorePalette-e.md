# AiTextOutlineRestorePalette-e

[![Direct](https://img.shields.io/badge/Direct%20Link-AiTextOutlineRestorePalette--e.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/outline/AiTextOutlineRestorePalette-e.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiTextOutlineRestorePalette-e.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Outline selected text (saving its attributes to the note) and restore it back, from a persistent palette
- Before outlining, character/paragraph attributes are serialized into the object's note as text
- On restore, the note is parsed to recreate the text frame, bringing back font, size, leading,
  kerning, fill color, alignment, horizontal/vertical scale, position, etc.
- Japanese-only typography attributes (orientation, kinsoku, mojikumi, tsume) are intentionally not handled in this English build
- The original outlines are moved to the outlined_text layer (dimmed and locked)
- The selected object's note is listed in the panel (Load button)
- Older notes without newer fields fall back to previous behavior (backward compatible)
- All DOM work is delegated to the main engine via BridgeTalk

### Main Features

- Outline with Memo: serialize properties into the note
- Restore Text: recreate the text from the note; the original outlines are moved to the outlined_text layer
- Show the selected object's note in the panel (Load button)

### Supported properties

- Text contents (multi-line)
- Font (PostScript name)
- Font size
- Leading
- Auto leading
- Horizontal & vertical scale
- Kerning method (metrics / optical / roman only / none)
- Proportional metrics (linked to the kerning method)
- Tracking
- Alignment (left / center / right / justify)
- Fill color (CMYK / RGB / Gray / Spot)
- Position (by geometricBounds; not shown in the list)

Note: Gradient/pattern fills and mixed attributes are not supported (the first character's value is used)

### Notes

https://note.com/dtp_tranist/n/nc476be8ad43c

### Script info

- Version: v2.0.0
