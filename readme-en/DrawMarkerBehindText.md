# Draw markers, underlines and strikethroughs behind text

[![Direct](https://img.shields.io/badge/Direct%20Link-DrawMarkerBehindText.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/DrawMarkerBehindText.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawMarkerBehindText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Draws a line behind the characters selected with the Type tool, or behind the selected text objects, fitted to the extent of their glyphs (the shape of the outlined characters).
- Choose the style (marker, underline, strikethrough, fill and more), color and weight in the dialog box, and check the result in the preview.
- The line is a path with live effects, so it can be adjusted later in the Appearance panel.

### Main Features

- Style
  - Underline: drawn just below the glyphs, with the top edge of the line on their bottom.
  - Marker: drawn centered on the bottom edge of the glyphs (default).
  - Straight + Shear: same position as Marker, with both ends slanted by the Free Distort effect. Set the slant with Angle (-80 to 80°; positive slants to the upper right).
  - Zigzag: same position as Marker, with the Zig Zag effect. Set Size (ridge height) and Ridges (count from end to end).
  - Fill: fills the glyph area (the line width is the glyph height).
  - Box: draws one box around the glyphs of each line.
  - Box per Character: draws a box around each character. Boxes are as tall as the line's glyphs and as wide as the glyph or as tall, whichever is larger (nearly square for kanji and kana).
  - For Box and Box per Character, set the space around the glyphs with Padding and the rounding with Corner Radius (0 for square corners).
  - Strikethrough: drawn through the vertical center of the glyphs.
- Color
  - Yellow (Y100), Red (M100Y100)
  - Text Color: uses the fill of the first target character. Set the line's opacity with Opacity next to it (30% by default).
  - Custom: click the swatch to choose a color with Illustrator's standard color picker.
- Weight: 40% of the glyph height by default, in the Stroke unit from Preferences.
- Vertical Offset: moves the line up or down from the style's position (positive is up, negative is down).
- Line: Solid, Dashed or Dotted. Dashed places dashes twice the line width, spaced by the line width; Dotted places round dots as large as the line width, every two line widths (Dotted always uses round caps).
- Cap: Butt or Round.
- White Text: when on, sets the fill of the target characters to white after drawing (also shown in the preview).
- Every style except Box and Box per Character gets the Outline Stroke effect. Those two group filled rectangles.
- Presets: save the current settings under a name and load them from the dropdown. Saved presets remain after Illustrator restarts.
  - Built-in presets: Yellow Marker, Slanted Highlighter, Red Underline, Red Wavy Line, Light Text-Color Fill, Box per Character (Pink), Rounded Box (Blue) and Text-Color Strikethrough. Lengths such as the weight follow the glyph height of the first target. Box per Character (Pink) and Rounded Box (Blue) turn on White Text. They cannot be deleted; a preset saved under the same name is used instead.
- Preview: while on, the selection is cleared and the line is shown. Closing the dialog box restores the original selection.

### Usage

1. Select characters with the Type tool (or select text objects)
2. Run the script
3. Choose the style, color, weight and so on (or load a preset)
4. Click OK

### Notes

- With text objects selected, one line is drawn for all characters of each object. Text inside groups is included; locked or hidden text is skipped.
- For multi-line text, one line is drawn for the combined extent of all lines, not one per line.
- The line is placed directly behind the original text.
- A selection of only spaces or line breaks has no glyphs, so nothing can be drawn.
- Text Color cannot be chosen when the first target character has no fill.
- Custom colors saved in presets are limited to RGB, CMYK and Grayscale.
- Weight, Vertical Offset, the zigzag Size, Padding and Corner Radius are fixed lengths, regardless of the type size.

### Original / Acknowledgements

Based on an idea by デザインゴリラ (Design Gorilla).  
https://note.com/gorilla_gg/n/naa26b079cb75

### Script info

- Version: v1.0.0

### Update History

- v1.0.0 (2026-11-01) Initial release
