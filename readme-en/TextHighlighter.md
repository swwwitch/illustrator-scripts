# Draw markers, underlines and strikethroughs on text

[![Direct](https://img.shields.io/badge/Direct%20Link-TextHighlighter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextHighlighter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextHighlighter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Draws lines such as markers, underlines and strikethroughs, or boxes, circles, hexagons and emphasis dots, on the characters selected with the Type tool or on the selected text objects, fitted to their glyphs (the shape of the outlined characters).
- They can go behind or in front of the text. Multi-line text is drawn line by line.
- Choose the style, colors, weight and more in the dialog box and check the result in the preview. The settings can be saved as presets.
- The lines and shapes are paths with live effects, so they can be adjusted later in the Appearance panel.

### Main Features

The dialog box has Presets on the left, and the Settings and Text panels on the right.

#### Style

Chosen in the dropdown at the top of Settings. The dropdown's tooltip describes the chosen style.

- Line styles
  - Underline: drawn just below the glyphs, with the top edge of the line on their bottom.
  - Marker: centered on the bottom edge of the glyphs (the default).
  - Straight + Shear: where Marker goes, with both ends slanted by the Free Distort effect (slanting keeps it centered).
  - Zigzag: where Marker goes, with the Zig Zag effect.
  - Hand-Drawn: where Marker goes, with the edges roughened by the Roughen effect.
  - Fill: fills the glyph area (the line width is the glyph height).
  - Strikethrough: through the vertical center of the glyphs.
  - Double Strikethrough: two lines through the vertical center of the glyphs.
- Shapes
  - Box: one box around the glyphs of each line (column in vertical text).
  - Circle: an ellipse inscribed in the Box area.
  - Hexagon: the Box area with 60° points added at the left and right (top and bottom in vertical text).
  - Per Character (right of the dropdown): draws a shape for each character. Shapes are as tall as the line's glyphs and as wide as the glyph or as tall, whichever is larger (nearly square for kanji and kana; swapped in vertical text). Circle becomes a circle whose diameter is the longer side, and Hexagon a regular hexagon whose flat-to-flat distance is the longer side. Punctuation, brackets and spaces are skipped.
- Emphasis Dots: a round dot above each character (right of it in vertical text). Punctuation, brackets and spaces get no dot.

#### Settings (left column)

- Stacking Order: Front or Back. Back (the default) draws behind the text; Front draws in front of it.
- Text: with White on, sets the fill of the target characters to white.
- Shape: turns Fill and Stroke on or off independently. Fill uses Fill Color; Stroke uses Stroke Color, Weight and Line Type.
- Fill Color: used for Fill, Emphasis Dots and the Fill of shapes. Choose White, Text Color or Custom (Text Color by default).
- Stroke Color: used for line styles and the Stroke of shapes. Choose Yellow (Y100), Red (M100Y100), Text Color or Custom (Yellow by default).
  - Text Color is the fill of the first target character; for Custom, click the swatch to choose with Illustrator's standard color picker.
  - The color not in use is dimmed.
- Opacity: opacity of the lines and shapes (100% by default).
- Weight: 40% of the glyph height by default, in the Stroke unit from Preferences.
- Offset: moves the line or shape up or down from the style's position (positive is up, negative is down; in vertical text it moves left or right, positive is left).
- Line Length: with Equalize on, for multi-line text, extends the lines and shapes of short lines to the length of the longest one (column length in vertical text). With Leading-Based on, lines, shapes and dots are positioned and sized by the leading instead of the glyph height of each line (width in vertical text), so every line gets the same height, centered on its characters (works with every style; Fill becomes as thick as the leading, and Per Character as large).

#### Settings (right column)

- Line Type: Solid, Dashed or Dotted.
  - Dashed: Segments dashes spaced by Gap, with the dash length fitted to the line so both ends finish with a dash. Blank Segments gives dashes about twice the line width.
  - Dotted: round dots as large as the line width, spaced close to Gap, with dots at both ends (Dotted always uses round caps).
  - Gap starts at the line width.
- Cap: Butt or Round.
- Angle: the slant of Straight + Shear (-80 to 80°; positive slants to the upper right).
- Size, Ridges, Points: the ridge height, the number of ridges from end to end, and Smooth (wavy) or Corner (zigzag) for Zigzag. Size is also the Hand-Drawn roughness.
- Padding (H), Padding (V): the left/right and top/bottom space around the glyphs for shapes, or the distance from the glyphs to the dots for Emphasis Dots (Padding (V) in horizontal text, Padding (H) in vertical text; negative values pull inward). While the link icon on the right is on, the two keep the same value (on by default).
- Corner Radius: the corner radius of Box (0 for square corners).
- Continuity: with Multi-Line on, makes a rounded Box spanning several lines look continuous. Only the left end of the first line and the right end of the last are rounded, and the line breaks keep square corners. With Stroke on, the stroke is left open on the edges where the lines continue (the left edge of every line but the first and the right edge of every line but the last; in vertical text, the top of the first column and the bottom of the last are rounded, and the top and bottom edges are left open).
- Rotate: turns each shape about its own center (-360 to 360°; positive is counterclockwise).
- Line Gap: the space between the two lines of Double Strikethrough (1/5 of the glyph height by default).
- Dot Size: the diameter of the emphasis dots.

Fields not in use are dimmed.

#### Text

Changes the formatting of the target characters. The preview reflects it, and the lines and shapes follow the reformatted glyphs.

- Kerning: No Change, 0, Auto or Optical. 0 turns auto kerning off and sets manual kerning to 0. 0 and Optical also turn Proportional Metrics off.
- Tracking: in 1/1000 em. Leave blank to keep it.
- Kerning (Before) / Kerning (After): manual kerning before the first target character and after the last one (1/1000 em). Leave blank to keep it.

#### Presets

- Click a preset in the list on the left to load it. The icons below the list save the current settings under a name or delete a preset. Saved presets remain after restarting Illustrator.
- Built-in presets: Underline, Dotted Underline, Wavy Line, Yellow Marker, Slanted Highlighter, Text-Color Strikethrough, Double Strikethrough, Rounded Box, Frame, Redaction, Box per Character, Circle per Character, Hexagon per Character and Emphasis Dots.
  - Except for Redaction, Underline and Box per Character, lengths such as the weight follow the glyph height of the first target.
  - Box per Character, Circle per Character, Hexagon per Character and Rounded Box turn on White under Text. Redaction uses Front for Stacking Order; the others use Back.
  - Built-in presets cannot be deleted; a preset saved under the same name is used instead.

#### Preview

While on, the selection is cleared so the lines and shapes are visible. Closing the dialog box restores the original selection.

#### Last state

The settings, the chosen preset and Preview at the time the dialog box closes (OK or Cancel) come back the next time it opens. They are kept until Illustrator quits; the first time after Illustrator starts, the Yellow Marker preset is chosen.

### Usage

1. Select characters with the Type tool (or select text objects)
2. Run the script
3. Pick a preset from the list on the left (or choose the style, colors, weight and more in Settings on the right)
4. Click OK

### Notes

- With text objects selected, all of their characters are targets. Text inside groups is included; locked and hidden text is skipped.
- Initial values such as the weight come from the glyph height of the first target's first line.
- In vertical text, line styles draw a vertical line for each column. Styles placed at the bottom edge of the glyphs, such as Underline and Marker, go to the right edge instead (the same side as Illustrator's underline). Straight + Shear slants the top and bottom ends.
- The lines and shapes are placed directly behind the original text (in front of it when Stacking Order is Front).
- Hide Edges (View > Hide Edges) is toggled while the script runs and toggled back at the end.
- Nothing can be drawn when only spaces or line breaks are selected, since they have no glyphs.
- Text Color cannot be chosen when the first target character has no fill.
- Only RGB, CMYK and Gray custom colors can be saved in presets.
- Weight, Offset, Gap, the zigzag Size, Padding (H), Padding (V), Corner Radius, Line Gap and Dot Size are fixed lengths, regardless of the type size.

### Original / Acknowledgements

The way the glyphs are measured is based on an article by デザインゴリラ (Design Gorilla).  
https://note.com/gorilla_gg/n/naa26b079cb75

### Script info

- Version: v1.3.1

### Update History

- v1.3.1 (2026-11-02) The dialog box now remembers its state at close (settings, chosen preset and Preview) and restores it the next time it opens, until Illustrator quits; the first time after Illustrator starts, the Yellow Marker preset is chosen
- v1.3.0 (2026-11-01) Added Leading-Based to Line Length (lines, shapes and dots are positioned and sized by the leading instead of the glyph height of each line, so every line gets the same height, centered on the line); moved Stacking Order, Text and Shape to the top of the left column and Line Length to its bottom, and Line Type, Segments, Gap and Cap to the top of the right column; Multi-Line now sits under the Continuity label and, with Stroke on, leaves the stroke open where the lines continue (a multi-line Frame looks continuous); split Padding into Padding (H) and Padding (V) with a link icon (presets from earlier versions load the same value into both); fixed the script sometimes stopping with an MRAP error while measuring the glyphs
- v1.2.1 (2026-11-01) Supports vertical text: line styles now draw a vertical line for each column instead of one horizontal line over the whole selection (styles placed at the bottom edge go to the right edge); renamed Vertical Offset to Offset, Equal Line Lengths to Equalize under Line Length, and White Text to White under Text; in vertical text, Offset now moves lines, shapes and emphasis dots left or right (positive is left); Equalize (Line Length) now also works for line styles in vertical text; fixed lines and shapes sometimes breaking at tate-chu-yoko numbers such as "90"
- v1.2.0 (2026-11-01) Reordered the built-in presets to follow Style; Added Equal Line Lengths (short lines get the length of the longest one); Added Segments (dashes per line; blank for automatic) and Gap (space between dashes) to Dashed, and Gap (space between dots) to Dotted; Added Multi-Line to Box (a rounded Box spanning several lines is rounded only at its start and end); The initial Line Gap of Double Strikethrough, and that of the preset of the same name, is now 1/5 of the glyph height; Hide Edges is toggled while the script runs and toggled back at the end; Line styles are now drawn per line in multi-line text (initial values such as the weight come from the first line); fixed a line being split in two when it starts with a small glyph such as "…" or "-"; renamed the dialog box to "Draw Lines and Shapes on Text", Auto Kerning to Kerning and its Metrics to Auto (as in Illustrator's Character panel); added tooltips to Stacking Order; renamed the presets Box per Character (Pink), Circle per Character (Pink) and Hexagon per Character (Blue) to Box per Character, Circle per Character and Hexagon per Character, and Rounded Box (Blue) to Rounded Box; Added Kerning (Before) and Kerning (After) (manual kerning) to Text; added Shape (Fill and Stroke, each on or off), and Rotate to Box, Circle and Hexagon, and merged Frame into a stroked Box (presets saved with Frame still load); split Color into Fill Color (White, Text Color, Custom) and Stroke Color (Yellow, Red, Text Color, Custom), with line styles using Stroke Color (presets from earlier versions still load); removed the Hand-Drawn Marker built-in preset; Wavy Line now has a Vertical Offset of -0.5 mm
- v1.1.0 (2026-11-01) Renamed the file from DrawMarkerBehindText.jsx to TextHighlighter.jsx (saved presets carry over); added Stacking Order (Front / Back) to Settings; Dashed now fits the dash length to the line so both ends finish with a dash, and Dotted fits the dot spacing so dots land on both ends; fixed the Outline Stroke effect not being applied (so Straight + Shear did not slant, for example) when characters were selected with the Type tool; Auto Kerning 0 and Optical now turn Proportional Metrics off; Dotted Underline now has a Vertical Offset of -0.6 mm, and Box per Character a Padding of -0.1 mm; Straight + Shear stays centered when slanted (the bottom edge shifts the opposite way to the top); Red Wavy Line is renamed Wavy Line and uses the text color
- v1.0.0 (2026-11-01) Initial release
