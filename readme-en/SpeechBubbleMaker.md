# Draw an underline or box with a speech-bubble tail around text

[![Direct](https://img.shields.io/badge/Direct%20Link-SpeechBubbleMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SpeechBubbleMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SpeechBubbleMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Draws an underline, a bracket or a box with a speech-bubble tail (a V, a slash or a taper) around the selected text. The line can go on any side, with color, cap, rounded corners and a shadow. Settings are previewed in a dialog box and can be saved as presets.

### Features

The dialog box has Presets centered at the top, the Tail and Text Options panels on the left below it, and the Line/Frame panel on the right. Preview sits at the left end of the button row.

#### Presets

- Picking a preset switches the tail and line/frame settings at once.
- Twelve built-in presets are included (simple speech bubble, wide curly, underline with a V, a V down right, slashes down right and left; bracket with a slash; rounded bracket with a V; box with a slash; rounded box with a slash; box with a V and shadow; curly bracket). Built-in presets size themselves to the font size of the first character.
- The save icon stores the current settings under a name; the trash icon deletes the selected saved preset (built-in presets cannot be deleted, and their names cannot be reused). Saved presets keep their sizes as saved and survive quitting Illustrator.
- Changing a setting by hand clears the preset selection.

#### Tail

- Shape: V bends the line into a tail (it stays a single path). Slash draws a slash out of the line. Taper narrows out of the line in curves to a point, like the middle of a curly bracket in AiCurlyBracketMaker.
- Direction: Down, Down right or Down left.
  - For a V or taper, Down puts the tip at the middle of the base; Down right puts it Width to the right of the base's left end; Down left puts it Width to the left of the base's right end.
  - A slash leans to the right or left at Angle (Down keeps it straight).
- Angle (slash pointing down right or left): measured from the line (default 45°, 5–90°; 90° is straight down). The slash's horizontal run follows Height and Angle.
- Root (slash pointing down right or left): Center starts the slash at the middle of the gap, apart from the line; Edge starts it at the edge of the gap, joined to the line.
- Position: from 0% (start of the line) to 100% (end), in the field or with the slider below it (hold Shift for 10% steps); default 50%. A tail that would run past the ends is pulled inward.
- Gap: how much of the line is cut away at the tail; for a V or taper, the width of its base. A slash with 0 keeps the line whole.
- Width (V or taper pointing down right or left): how far the tip moves sideways.
- Height: how far the tail reaches out from the line.

#### Line/Frame

- Shape: Underline draws a line below the glyphs. Bracket turns both ends of the line up toward the text. Box encloses the glyphs in a rectangle. A bracket whose Corner radius equals its Ticks gets curly-bracket ends.
- Side: Top, Bottom, Left or Right, laid out around a cross. Each time the dialog box opens, Bottom is picked when the first text is horizontal and Right when it is vertical (the same happens when picking a preset that has no side of its own).
  - On Top, the tail points up.
  - On Left or Right the line starts at the top, Down right leans toward the bottom and Down left toward the top.
- Distance: from the glyphs to the line; the top and bottom padding of a box.
- Extension: how far the line runs past the glyphs (negative values shorten it); the side padding of a box.
- Ticks (bracket): length of the turned-up ends.
- Corner radius (bracket or box): 0 keeps the corners sharp.
- Shadow (box): adds a second line this far out on the right and bottom (0 for none). Its ends line up with where the box's corner rounding starts; the tail's right edge is offset in parallel and stops at the tip's height, and the bottom line stops short of the tail's left edge.
- Weight: stroke weight of the line and tail.
- Color: Text (default) uses the fill of each text's first character (black when it has no fill); the other radio button (custom, shown with just a swatch) uses the swatch's color. Clicking the swatch opens Illustrator's Color Picker and switches to custom.
- Cap: Butt (default) or Round. Butt keeps the corners pointed; Round rounds the corners too.

#### Text Options

- Use the frame of area type (on by default; available when area type is selected): area type is drawn along its frame instead of its glyphs (Distance and Extension count as 0; the two fields are dimmed when every selected text is area type). Point type and type on a path are still measured by their glyphs. Saved in presets too.
- Center as area type (box only): turns the text into area type the size of the box, centered horizontally and vertically. Point type is converted; area type gets its frame resized and moved (type on a path is left alone).
- Group with text: groups the line and tail together with the text. When off, the line and tail are grouped on their own, right in front of the text.

#### Other

- Preview (left end of the button row): shows the result on the canvas (on by default).
- While the script runs, Hide Edges and Hide Bounding Box are toggled to make the preview easier to see, and toggled back at the end.
- Lengths use the General unit under Preferences > Units; the weight uses the Stroke unit.
- The first time, the sizes are derived from the font size of the first character. The settings used with OK (custom color included) are remembered until Illustrator quits and come back next time (only Side is picked again from the text each time).

### Usage

1. Select text (several objects are fine, including text inside groups).
2. Run the script, pick a preset or set the tail and line/frame shapes and sizes, and click OK.

### Notes

- One set of lines is drawn around the bounds of all the glyphs of the text (the frame, when area type is measured by its frame). With several lines, an underline goes below the last line (outside the outermost column on the left or right).
- The area type conversion and vertical centering run after OK, so the preview does not show them.
- Locked or hidden text, and text containing only spaces, is skipped.

### Update History

- v1.0.0 (2026-11-02): Initial release
- v1.0.1 (2026-11-02): Added the built-in preset Simple speech bubble and grouped the built-in presets with separators (staple, underlines, brackets, boxes). Renamed the Options panel to Text Options and Center the text as area type to Center as area type. Moved Shadow below Corner radius and made the save icon slightly smaller. The position slider moves in 10% steps with Shift held. Fixed the Color Picker sometimes opening in grayscale
- v1.0.2 (2026-11-02): Added the built-in preset Wide curly (an underline with a taper and round caps)
