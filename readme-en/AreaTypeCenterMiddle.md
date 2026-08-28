# Center Area Type both vertically and horizontally

[![Direct](https://img.shields.io/badge/Direct%20Link-AreaTypeCenterMiddle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AreaTypeCenterMiddle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AreaTypeCenterMiddle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Sets the vertical alignment (center) and the justification (center) of the selected Area Type frames in one pass.

Selected closed paths are converted to Area Type first and then treated the same way, so a rectangle can become a finished button or label in a single run.

### Features

- Center the selected Area Type frames both vertically and horizontally
- Convert closed paths (rectangles and the like) to Area Type filled with sample text
- With one rectangle and one text object selected, pour that text into the shape
- Process any number of objects at once

### Usage

1. Select Area Type frames or closed paths.
2. Run the script.

There is no dialog. The resulting Area Type frames are left selected.

### Behavior by selection

**Area Type**

Sets the vertical alignment and the justification to center.

**Closed paths (rectangle, ellipse, compound path, …)**

Converts them to Area Type, pours in the sample text (“Typography”, or 「山路を登りながら」 in a Japanese locale), then centers it. The fill and stroke of the path are removed as part of the conversion.

**One closed path plus one text object**

Pours the contents of the selected text instead of the sample text. Font, size and fill color are carried over from the original, and the original text is removed once poured. Point Type and Type on a Path both work as the source.

### User settings

The sample text and its formatting can be changed in the "User settings" block at the top of the script.

- `DUMMY_TEXT_JA` / `DUMMY_TEXT_EN`: the sample text to pour (defaults to 「山路を登りながら」 / “Typography”)
- `DUMMY_FONT_JA` / `DUMMY_FONT_EN`: preferred font names, tried in order until one is found
- `DUMMY_FONT_SIZE`: font size of the sample text (defaults to 10)

### Notes

- Vertical alignment has no DOM API, so a temporary .aia action is generated, loaded and played internally, then unloaded.
- For a compound path only the first path becomes the Area Type frame; the emptied compound path is removed.
- The "one closed path plus one text" case is detected only when the selection holds exactly two objects. With three or more selected, the paths get the sample text.
- Only font, size and fill color are carried over; other character attributes fall back to the defaults.

### Update History

- v1.0.0 (2026-08-28): Initial release
