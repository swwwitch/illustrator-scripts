# ConvertFontInfo

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertFontInfo.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ConvertFontInfo.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertFontInfo.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An Illustrator script that converts selected text into font information.
- Choose a conversion format in the dialog and rewrite the target text with an immediate preview.
- The right column shows conversion results generated from the actual font information of the selected text.
- Reapplies the selected format on OK, and restores original contents, font, size, and justification when canceled.

### Main Features

- Choose from font family, style, font family + style, PostScript name, full name + size, or detail view
- Defaults to Font Family + Style
- Switch conversion formats with the F/S/B/P/M/D keys
- Confirm with the Enter / Return key (OK is the default button)
- Detail view separates label lines and value lines, and sets justification to left
- Keeps references to the selected text frames, so the same targets are updated even if the selection changes during preview
- Converts font size according to the Illustrator "Type Units" preference and rounds it to two decimals
- Japanese and English UI support

### Update History

- v1.0.0 (20250509): Initial version
- v1.1.0 (20260428): Expanded conversion formats and added actual font preview, keyboard shortcuts, and final reapply on OK
- v1.1.1 (20260608): Allow running from a partial character selection (TextRange), made OK the default button, internal refactor
- v1.1.2 (20260611): Removed the redundant re-apply on OK to avoid crashes (commit the previewed state as-is)

### Script info

- Version: v1.1.2
