# Draw markers, underlines and strikethroughs on text

[![Direct](https://img.shields.io/badge/Direct%20Link-TextHighlighter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextHighlighter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextHighlighter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Draws lines such as markers, underlines and strikethroughs, or boxes, circles, hexagons and emphasis dots, on the characters selected with the Type tool or on the selected text objects, fitted to their glyphs (the shape of the outlined characters).
- They can go behind or in front of the text. Multi-line text is drawn line by line, and the line starts, lengths and leading can be aligned.
- Choose the style, colors, weight and more in the dialog box and check the result in the preview. The settings can be saved as presets.
- The color, kerning and tracking of the target characters can be changed at the same time. For shapes drawn per line, the kerning before and after can open up automatically to match the shape's padding (with manual fine-tuning).
- The lines and shapes are paths with live effects, so they can be adjusted later in the Appearance panel.

### Main Features

The dialog box has Presets on the left, and the Settings and Text panels on the right.

#### Style

Chosen in the dropdown at the top of Settings. The dropdown's tooltip describes the chosen style.

- Line styles
  - Underline: drawn just below the glyphs, with the top edge of the line on their bottom.
  - Marker: centered on the bottom edge of the glyphs (the default).
  - Straight + Shear: where Marker goes, with both ends slanted by the Free Distort effect (slanting keeps it centered).
  - Zigzag: where Marker goes; points are added with the Roughen effect (size 0), then the Zig Zag effect (0 ridges) is applied, so the ridge pitch stays the same whatever the line length.
  - Fill: fills the glyph area (the line width is the glyph height).
  - Strikethrough: through the vertical center of the glyphs.
  - Double Strikethrough: two lines through the vertical center of the glyphs.
- Shapes
  - Box: one box around the glyphs of each line (column in vertical text).
  - Circle: an ellipse inscribed in the Box area.
  - Hexagon: the Box area with 60° points added at the left and right (top and bottom in vertical text).
  - Unit (right column under Settings): Line (the default) draws one shape per line (column in vertical text); Per Character draws a shape for each character. Shapes are as tall as the line's glyphs and as wide as the glyph or as tall, whichever is larger (nearly square for kanji and kana; swapped in vertical text). Circle becomes a circle whose diameter is the longer side, and Hexagon a regular hexagon whose flat-to-flat distance is the longer side. Punctuation, brackets and spaces are skipped.
- Emphasis Dots: a round dot above each character (right of it in vertical text). Punctuation, brackets and spaces get no dot.

#### Settings (Fill and Stroke, spanning both columns)

- Stacking Order (right of Style): Front or Back. Back (the default) draws behind the text; Front draws in front of it.
- Fill: the color for Fill, Emphasis Dots and the fill of shapes. Choose Custom, Text Color or White (Text Color by default). With Draw at the start of the row on, shapes (Box, Circle, Hexagon) are filled (on by default).
- Stroke: the color for line styles and the stroke of shapes. Choose Custom, Text Color, Yellow (Y100) or Red (M100Y100) (Yellow by default). With Draw at the start of the row on, shapes get a stroke (using Weight and Line Type; off by default).
  - Text Color is the fill of the first target character; for Custom, click the swatch to choose with Illustrator's standard color picker.
  - The color not in use is dimmed. Draw is available for shapes only.

#### Settings (left column: lines and effects)

- Weight: 40% of the glyph height by default, in the Stroke unit from Preferences.
- Line Type: Solid, Dashed or Dotted.
  - Dashed: Segments dashes spaced by Gap, with the dash length fitted to the line so both ends finish with a dash. Blank Segments gives dashes about twice the line width.
  - Dotted: round dots as large as the line width, spaced close to Gap, with dots at both ends (Dotted always uses round caps).
  - Gap starts at the line width.
- Cap: Butt or Round.
- Line Gap: the space between the two lines of Double Strikethrough (1/5 of the glyph height by default).
- Angle: the slant of Straight + Shear (-80 to 80°; positive slants to the upper right).
- Size, Ridges, Points: the ridge height, the number of ridges per inch (25.4 mm) of the line, and Smooth (wavy) or Corner (zigzag) for Zigzag. Ridges is also the Roughen detail for Hand-Drawn (Zigzag and Hand-Drawn together share the number).
- Roughen: with Hand-Drawn on, the Roughen effect (corner points) is applied to the lines and shapes so their edges wobble like hand drawing (line styles get it on the center line, before Outline Stroke). Size is a percentage of the line or shape (initially 1%); the detail uses Ridges (per inch). Works with every style except Emphasis Dots.

#### Settings (right column: placement and size)

- Opacity: opacity of the lines and shapes (100% by default).
- Offset: moves the line or shape up or down from the style's position (positive is up, negative is down; in vertical text it moves left or right, positive is left).
- Padding (H), Padding (V): the left/right and top/bottom space around the glyphs for shapes, or the distance from the glyphs to the dots for Emphasis Dots (Padding (V) in horizontal text, Padding (H) in vertical text; negative values pull inward). While the link icon on the right is on, the two keep the same value (on by default).
- Line Length: how the lines and shapes of multi-line text are aligned.
  - Align Starts: starts the lines and shapes where the earliest line starts (the top of the column in vertical text). The ends stay as they are.
  - Align Both Ends: extends the lines and shapes of short lines to the length of the longest one (column length in vertical text).
  - Only one of Align Starts and Align Both Ends can be on. They work with line styles and with shapes when Unit is Line.
- Line Height: with Leading-Based on, positions and sizes lines, shapes and dots by the leading instead of the glyph height of each line (width in vertical text). Every line gets the same height, centered on its characters, and consecutive lines are spaced exactly one leading apart (spacing restarts after paragraph space or blank lines). Works with every style; Fill becomes as thick as the leading, and Per Character as large.
- Corner Radius: the corner radius of Box (0 for square corners).
- Continuity: with Multi-Line on, makes a rounded Box spanning several lines look continuous. Only the left end of the first line and the right end of the last are rounded, and the line breaks keep square corners. With Stroke on, the stroke is left open on the edges where the lines continue (the left edge of every line but the first and the right edge of every line but the last; in vertical text, the top of the first column and the bottom of the last are rounded, and the top and bottom edges are left open).
- Rotate: turns each shape about its own center (-360 to 360°; positive is counterclockwise).
- Dot Size: the diameter of the emphasis dots.

Fields not in use are dimmed.

#### Text

Changes the formatting of the target characters. The preview reflects it, and the lines and shapes follow the reformatted glyphs.

- Text Color: sets the fill of the target characters after drawing. Choose Custom, No Change (default) or White. For Custom, click the swatch and pick a color in Illustrator's standard Color Picker.
- Kerning: No Change, 0, Auto or Optical. 0 turns auto kerning off and sets manual kerning to 0. 0 and Optical also turn Proportional Metrics off.
- Kerning (Before) / Kerning (After): manual kerning before the first target character and after the last one (1/1000 em). Leave blank to keep it.
- Tracking: in 1/1000 em. Leave blank to keep it.
- Auto-Adjust Kerning to Padding: when Box, Circle or Hexagon is drawn per line, sets Kerning (Before) / (After) to Padding (H) / type size x 1000 to make room for the shape (on by default; Padding (V) in vertical text; uses the size of the first and last target characters, and leaves the kerning unchanged when the padding is 0 or less). While on, the Kerning (Before) / (After) fields show the values computed for the first target. Editing a field keeps the difference from the computed value as a fine adjustment, which stays applied when the padding changes (clear the field to drop it). Adjustments are saved with presets.

#### Presets

- Click a preset in the list on the left to load it. Below the list, Delete sits at the left, and Save (saves the current settings under a name) and Overwrite (writes the current settings into the selected preset without asking for a name or confirmation) at the right. Saved presets remain after restarting Illustrator.
- Built-in presets (in Style order): Underline, Dotted Underline, Yellow Marker, Yellow Marker (Hand-Drawn), Slanted Highlighter, Wavy Line, Zigzag, Text-Color Strikethrough, Double Strikethrough, Rounded Box, Band per Line, Band per Line (Touching), Band per Line (Hand-Drawn), Frame, Frame per Line, Frame per Line (Hand-Drawn), Redaction, Box per Character, Circle per Character, Hexagon per Character and Emphasis Dots.
  - Text-Color Strikethrough and Double Strikethrough take lengths such as the weight from the glyph height of the first target; the others use fixed lengths.
  - Rounded Box, Band per Line, Band per Line (Touching), Band per Line (Hand-Drawn), Box per Character, Circle per Character and Hexagon per Character set Text Color to White. All use Back for Stacking Order.
  - Band per Line draws a box for each line with small gaps, looking continuous (Multi-Line on); Band per Line (Touching) leaves no gaps, with Multi-Line off. Both turn on Align Starts and Leading-Based.
  - Frame per Line draws Frame for each line, with Align Starts on and Multi-Line off.
  - The Hand-Drawn presets turn on Hand-Drawn under Roughen (Yellow Marker (Hand-Drawn): Size 1%, Detail 10; Band per Line (Hand-Drawn): 0.7%, 3; Frame per Line (Hand-Drawn): 0.4%, 10).
  - The three per-character presets set Kerning to 0 and Tracking to 170 under Text; the presets that draw shapes per line turn on Auto-Adjust Kerning to Padding.
  - Every preset carries the Line Length, Line Height, Roughen and Text settings, so choosing another preset switches them too.
  - Changing a setting deselects the list, so clicking the same preset again brings its settings back. After that, Overwrite still writes into the preset chosen just before (its name appears in the tooltip), and Save offers that name too.
  - Overwriting a built-in preset makes a saved preset of the same name; deleting it brings back the original settings.
  - Built-in presets cannot be deleted; a preset saved under the same name is used instead.

#### Preview and Reset

While the dialog box is open, the selection is cleared and the lines and shapes are always previewed. Closing the dialog box restores the original selection.
Reset returns to the settings and the preset chosen when the dialog box opened, with Kerning (Before) / (After) and Tracking cleared and Auto-Adjust Kerning to Padding off.

#### Last state

The settings and the chosen preset at the time the dialog box closes (OK or Cancel) come back the next time it opens. They are kept until Illustrator quits; the first time after Illustrator starts, the Yellow Marker preset is chosen.

### Usage

1. Select characters with the Type tool (or select text objects)
2. Run the script
3. Pick a preset from the list on the left (or choose the style, colors, weight and more in Settings, and the text color, kerning and more in Text, on the right)
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

### Article

https://note.com/dtp_tranist/n/na1804afb7fe2

### Script info

- Version: v1.7.4

### Update History

- v1.7.4 (2026-11-02) Renamed Match Padding to Auto-Adjust Kerning to Padding
- v1.7.3 (2026-11-02) Added a tooltip to Auto under Kerning and expanded the Optical tooltip; internal cleanup (clearer names, shared duplicate code, split long functions)
- v1.7.2 (2026-11-02) Merged Fill and Stroke under Shape into a Draw checkbox at the start of the Fill and Stroke rows (formerly Fill Color and Stroke Color); Stroke now lists Custom, Text Color, Yellow, then Red; renamed Even under Leading to Leading-Based under Line Height; updated the overview and README
- v1.7.1 (2026-11-02) Fill Color now lists Custom, Text Color, then White
- v1.7.0 (2026-11-02) Turned the White checkbox under Text into Text Color radio buttons (No Change, White, Custom); Custom takes its color from the Color Picker via the swatch (White in presets from earlier versions loads as White when on and No Change when off); added Match Padding to Text (for Box, Circle and Hexagon per line, Kerning (Before) / (After) come from Padding (H) and the type size, with fine adjustment in the fields; on by default, including presets from earlier versions); the preview is now always on, replacing the Preview checkbox with Reset (returns to the state at opening, clearing Kerning (Before) / (After) and Tracking and turning Match Padding off); preset icons now show Delete on the left and Save / Overwrite on the right; Shape, Fill Color and Stroke Color moved below Style as rows spanning both Settings columns, with their radio buttons on one line (Custom and its swatch first); regrouped the two Settings columns into lines and per-style effects on the left (Weight, Line Type, Segments, Gap, Cap, Line Gap, Angle, Zigzag, Roughen) and placement and size on the right (Opacity, Offset, Padding, Line Length, Leading, Unit, Corner Radius, Continuity, Rotate, Dot Size); moved Text Color to the top of the Text panel, laid the Kerning radio buttons out in a row, and turned Kerning (Before), Tracking and Kerning (After) into three fields with their labels above
- v1.6.0 (2026-11-02) Merged the Roughen detail into Ridges, moved Roughen below the Zigzag fields in the right column and Padding (H) / Padding (V) to the bottom of the left column; turned the Per Character checkbox into Unit radio buttons (Line / Per Character); moved Stacking Order to the right of Style and Unit to the bottom right of the right column under Settings
- v1.5.0 (2026-11-02) Added an Overwrite icon to Presets (writes the current settings into the selected preset without asking for a name or confirmation); turned the Hand-Drawn style into a Hand-Drawn option under Roughen that works with every style (with Size in % and Detail; the former Hand-Drawn style loads as Marker with Hand-Drawn on); retuned the built-in presets, added Yellow Marker (Hand-Drawn), Band per Line (Touching), Band per Line (Hand-Drawn), Frame per Line and Frame per Line (Hand-Drawn), and reordered them by Style (fixed lengths except Text-Color Strikethrough and Double Strikethrough; Redaction now uses Back; every preset carries the Line Length, Leading and Text settings); fixed clicking the chosen preset again after changing settings not loading it, so the saved Offset, Weight and so on could not be brought back (changing a setting now deselects the list)
- v1.4.0 (2026-11-02) Added Align Starts to Line Length (for multi-line text, only the starts of the lines and shapes are aligned to the earliest line; the top of the column in vertical text); renamed Equalize to Align Both Ends and Leading-Based to Even under its own Leading row; added the Zigzag and Band per Line built-in presets (Band per Line: a box for each line, with aligned starts and even leading); Zigzag (Wavy Line) now adds points with the Roughen effect before the Zig Zag effect, and Ridges is now per inch instead of end to end (short and long lines get the same ridge pitch; initial value 12); fixed Even (Leading) sometimes shifting the shapes of kana-heavy or Latin lines slightly up or down (left or right in vertical text), leaving uneven gaps between lines
- v1.3.1 (2026-11-02) The dialog box now remembers its state at close (settings, chosen preset and Preview) and restores it the next time it opens, until Illustrator quits; the first time after Illustrator starts, the Yellow Marker preset is chosen
- v1.3.0 (2026-11-01) Added Leading-Based to Line Length (lines, shapes and dots are positioned and sized by the leading instead of the glyph height of each line, so every line gets the same height, centered on the line); moved Stacking Order, Text and Shape to the top of the left column and Line Length to its bottom, and Line Type, Segments, Gap and Cap to the top of the right column; Multi-Line now sits under the Continuity label and, with Stroke on, leaves the stroke open where the lines continue (a multi-line Frame looks continuous); split Padding into Padding (H) and Padding (V) with a link icon (presets from earlier versions load the same value into both); fixed the script sometimes stopping with an MRAP error while measuring the glyphs
- v1.2.1 (2026-11-01) Supports vertical text: line styles now draw a vertical line for each column instead of one horizontal line over the whole selection (styles placed at the bottom edge go to the right edge); renamed Vertical Offset to Offset, Equal Line Lengths to Equalize under Line Length, and White Text to White under Text; in vertical text, Offset now moves lines, shapes and emphasis dots left or right (positive is left); Equalize (Line Length) now also works for line styles in vertical text; fixed lines and shapes sometimes breaking at tate-chu-yoko numbers such as "90"
- v1.2.0 (2026-11-01) Reordered the built-in presets to follow Style; Added Equal Line Lengths (short lines get the length of the longest one); Added Segments (dashes per line; blank for automatic) and Gap (space between dashes) to Dashed, and Gap (space between dots) to Dotted; Added Multi-Line to Box (a rounded Box spanning several lines is rounded only at its start and end); The initial Line Gap of Double Strikethrough, and that of the preset of the same name, is now 1/5 of the glyph height; Hide Edges is toggled while the script runs and toggled back at the end; Line styles are now drawn per line in multi-line text (initial values such as the weight come from the first line); fixed a line being split in two when it starts with a small glyph such as "…" or "-"; renamed the dialog box to "Draw Lines and Shapes on Text", Auto Kerning to Kerning and its Metrics to Auto (as in Illustrator's Character panel); added tooltips to Stacking Order; renamed the presets Box per Character (Pink), Circle per Character (Pink) and Hexagon per Character (Blue) to Box per Character, Circle per Character and Hexagon per Character, and Rounded Box (Blue) to Rounded Box; Added Kerning (Before) and Kerning (After) (manual kerning) to Text; added Shape (Fill and Stroke, each on or off), and Rotate to Box, Circle and Hexagon, and merged Frame into a stroked Box (presets saved with Frame still load); split Color into Fill Color (White, Text Color, Custom) and Stroke Color (Yellow, Red, Text Color, Custom), with line styles using Stroke Color (presets from earlier versions still load); removed the Hand-Drawn Marker built-in preset; Wavy Line now has a Vertical Offset of -0.5 mm
- v1.1.0 (2026-11-01) Renamed the file from DrawMarkerBehindText.jsx to TextHighlighter.jsx (saved presets carry over); added Stacking Order (Front / Back) to Settings; Dashed now fits the dash length to the line so both ends finish with a dash, and Dotted fits the dot spacing so dots land on both ends; fixed the Outline Stroke effect not being applied (so Straight + Shear did not slant, for example) when characters were selected with the Type tool; Auto Kerning 0 and Optical now turn Proportional Metrics off; Dotted Underline now has a Vertical Offset of -0.6 mm, and Box per Character a Padding of -0.1 mm; Straight + Shear stays centered when slanted (the bottom edge shifts the opposite way to the top); Red Wavy Line is renamed Wavy Line and uses the text color
- v1.0.0 (2026-11-01) Initial release
