# Unify mixed font sizes to the first character's size

[![Direct](https://img.shields.io/badge/Direct%20Link-FontSizeToScaleConverter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/FontSizeToScaleConverter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontSizeToScaleConverter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Sometimes only part of a text has a different font size: larger digits in a heading, or a smaller unit after a number. Text like this is awkward to handle later, when you change the size of the whole text or apply a character style.

This script sets every character in a text to the font size of its first character. The size difference is converted into horizontal and vertical scale, so the characters look the same size as before.

For example, if a text starts at 10 pt and contains a 15 pt character, that character becomes 10 pt with 150% horizontal scale and 150% vertical scale.

### Features

- Sets the other characters in each text to the font size of its first character
- Converts the size difference into horizontal and vertical scale to keep the visual size. Existing scale is multiplied, not replaced
- For characters with auto leading, fixes the current leading as a value so the line spacing does not change
- Works with point type, area type, and type on a path
- Processes several text objects, and text inside groups, in one run

### How to use

1. Select text objects with mixed font sizes (multiple objects and groups are fine)
2. Run the script

There is no dialog; the conversion runs immediately.

### Options

These can be changed in the user settings at the top of the script:

| Setting | Description |
|---|---|
| `PRESERVE_LEADING` | Set to `false` to leave auto leading as it is instead of fixing it (default `true`) |
| `FONT_SIZE_TOLERANCE` | Largest difference treated as the same size, in pt (default `0.001`) |

### Notes

- The base is the first character of each text. Sizes are not matched across different text objects
- Characters whose scale would fall outside 1%–10000% are left unchanged, because their look cannot be kept
- Characters that had auto leading end up with a fixed leading. It no longer follows later font size changes
- Locked and hidden objects are skipped

### Version history

- v1.0.0 (2026-06-18): Initial version
