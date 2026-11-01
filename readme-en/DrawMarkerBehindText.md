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
  - Zigzag: same position as Marker, with the Zig Zag effect. Set Size (ridge height), Ridges (count from end to end) and Points (Smooth for a wavy line, Corner for sharp ridges).
  - Hand-Drawn: same position as Marker, with the edges roughened by the Roughen effect. Set the amount with Size.
  - Fill: fills the glyph area (the line width is the glyph height).
  - Strikethrough: drawn through the vertical center of the glyphs.
  - Double Strikethrough: two lines through the vertical center of the glyphs, spaced by Line Gap.
  - Box: draws one box around the glyphs of each line (column in vertical text).
  - Frame: the Box area as a stroked outline with no fill, using Weight and Line Type.
  - Circle: an ellipse inscribed in the Box area.
  - Hexagon: the Box area with 60° points added on the left and right (top and bottom in vertical text).
  - For Box, Frame, Circle and Hexagon, set the space around the glyphs with Padding (negative values pull inward); Box and Frame take Corner Radius (0 for square corners).
  - Per Character: draws Box, Frame, Circle or Hexagon for each character. Shapes are as tall as the line's glyphs and as wide as the glyph or as tall, whichever is larger (nearly square for kanji and kana; swapped in vertical text); Circle becomes a circle whose diameter is the longer side, and Hexagon a regular hexagon whose flat-to-flat width is the longer side. Punctuation, brackets and spaces are skipped, like Emphasis Dots.
  - Emphasis Dots: a round dot above each character (right of it in vertical text). Set the diameter with Dot Size and the distance from the glyphs with Padding. Punctuation, brackets and spaces get no dot.
- Color
  - Yellow (Y100), Red (M100Y100)
  - Text Color: uses the fill of the first target character.
  - Custom: click the swatch to choose a color with Illustrator's standard color picker.
- Opacity: opacity of the lines and boxes (100% by default), for any color.
- Weight: 40% of the glyph height by default, in the Stroke unit from Preferences.
- Vertical Offset: moves the line up or down from the style's position (positive is up, negative is down).
- Line Type: Solid, Dashed or Dotted. Dashed places dashes twice the line width, spaced by the line width; Dotted places round dots as large as the line width, every two line widths (Dotted always uses round caps).
- Cap: Butt or Round.
- Text (formatting of the target characters, also shown in the preview)
  - Auto Kerning: No Change, 0, Metrics or Optical. 0 turns auto kerning off and sets manual kerning to 0.
  - Tracking: in 1/1000 em. Leave blank to keep it.
  - White Text: when on, sets the fill of the target characters to white.
  - Lines and shapes are drawn on the glyphs after the formatting changes.
- Line styles get the Outline Stroke effect. Box groups filled rectangles, Circle groups filled circles, Hexagon groups filled hexagons, Frame groups stroked rectangles, and Emphasis Dots groups filled circles.
- Presets: click a preset in the list on the left of the dialog box to load it. Save the current settings under a name, or delete a preset, with the icons below the list. Saved presets remain after Illustrator restarts.
  - Built-in presets: Redaction, Underline, Dotted Underline, Yellow Marker, Slanted Highlighter, Red Wavy Line, Box per Character (Pink), Circle per Character (Pink), Hexagon per Character (Blue), Rounded Box (Blue), Hand-Drawn Marker, Frame, Emphasis Dots, Double Strikethrough and Text-Color Strikethrough. Except for Redaction and Underline, lengths such as the weight follow the glyph height of the first target. Box per Character (Pink), Circle per Character (Pink), Hexagon per Character (Blue) and Rounded Box (Blue) turn on White Text. They cannot be deleted; a preset saved under the same name is used instead.
- Preview: while on, the selection is cleared and the line is shown. Closing the dialog box restores the original selection.

### Usage

1. Select characters with the Type tool (or select text objects)
2. Run the script
3. Pick a preset in the list on the left (or choose the style, color, weight and so on under Settings)
4. Click OK

### Notes

- With text objects selected, one line is drawn for all characters of each object. Text inside groups is included; locked or hidden text is skipped.
- For multi-line text, line styles draw one line for the combined extent of all lines, not one per line. Boxes and Emphasis Dots are drawn per line.
- Line styles are meant for horizontal text; in vertical text the lines stay horizontal.
- The line is placed directly behind the original text.
- A selection of only spaces or line breaks has no glyphs, so nothing can be drawn.
- Text Color cannot be chosen when the first target character has no fill.
- Custom colors saved in presets are limited to RGB, CMYK and Grayscale.
- Weight, Vertical Offset, the zigzag Size, Padding, Corner Radius, Line Gap and Dot Size are fixed lengths, regardless of the type size.

### Original / Acknowledgements

Based on an idea by デザインゴリラ (Design Gorilla).  
https://note.com/gorilla_gg/n/naa26b079cb75

### Script info

- Version: v1.0.0

### Update History

- v1.0.0 (2026-11-01) Initial release
